import { useEffect, useMemo, useState } from "react";
import axios from "axios";
import * as XLSX from "xlsx";
import {
  Activity, AlertCircle, ArrowDownRight, ArrowUpRight, Award, BarChart3,
  Bell, BookOpen, BrainCircuit, CalendarDays, Check, ChevronDown, Clock3,
  Download, Eye, EyeOff, FileSpreadsheet, Filter, GraduationCap, LayoutDashboard, LineChart,
  LockKeyhole, LogOut, Mail, Menu, Moon, Pencil, Plus, Search, Settings2, ShieldCheck, Sparkles,
  Sun, Target, Trash2, TrendingUp, UserRound, Users, X, Zap
} from "lucide-react";
import {
  Bar, BarChart, CartesianGrid, Cell, Legend, Pie, PieChart, ResponsiveContainer,
  Tooltip, XAxis, YAxis
} from "recharts";
import "./App.css";

const API = "http://127.0.0.1:8000/api/predict";
const STUDENT_API = "http://127.0.0.1:8000/api/students/";
const HISTORY_API = "http://127.0.0.1:8000/api/history/";
const AUTH_API = "http://127.0.0.1:8000/api/auth";
const AUTH_TOKEN_KEY = "edupredict-admin-token";

// Attach the current admin token to protected FastAPI requests.
axios.interceptors.request.use(config => {
  const token = localStorage.getItem(AUTH_TOKEN_KEY);
  if (token) {
    config.headers = config.headers || {};
    config.headers.Authorization = `Bearer ${token}`;
  }
  return config;
});

const initialStudent = {
  name: "", email: "", course: "BCA AI & DS", attendance: 80,
  study_hours: 5, previous_marks: 70, assignment_score: 75, sleep_hours: 7
};

const performanceColors = {
  Excellent: "#16a37a",
  Average: "#e6a23c",
  "At Risk": "#e05b67"
};

const BULK_COLUMN_ALIASES = {
  name: ["name", "studentname", "fullname", "studentfullname"],
  email: ["email", "emailaddress", "studentemail"],
  course: ["course", "program", "programme", "department"],
  attendance: ["attendance", "attendancepercent", "attendancepercentage", "attendancepct"],
  study_hours: ["studyhours", "studyhoursday", "dailystudyhours", "studyhoursperday"],
  previous_marks: ["previousmarks", "previousmarkspercent", "previousmarkspercentage", "marks", "previousscore"],
  assignment_score: ["assignmentscore", "assignmentscorepercent", "assignmentmarks"],
  sleep_hours: ["sleephours", "sleephoursday", "dailysleephours"]
};

const REQUIRED_BULK_COLUMNS = [
  "name", "course", "attendance", "study_hours",
  "previous_marks", "assignment_score", "sleep_hours"
];

function normaliseExcelHeader(value) {
  return String(value ?? "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");
}

function downloadExcelTemplate() {
  const headers = [[
    "name", "email", "course", "attendance",
    "study_hours", "previous_marks", "assignment_score", "sleep_hours"
  ]];
  const templateSheet = XLSX.utils.aoa_to_sheet(headers);
  templateSheet["!cols"] = [
    { wch: 24 }, { wch: 30 }, { wch: 20 }, { wch: 16 },
    { wch: 18 }, { wch: 19 }, { wch: 20 }, { wch: 16 }
  ];
  const instructionsSheet = XLSX.utils.aoa_to_sheet([
    ["EDUPREDICT — BULK STUDENT IMPORT"],
    [],
    ["Column", "Instructions"],
    ["name", "Required. Student full name."],
    ["email", "Optional, but recommended for duplicate detection."],
    ["course", "Required. Example: BCA AI & DS."],
    ["attendance", "Required. Number from 0 to 100."],
    ["study_hours", "Required. Average study hours per day, from 0 to 24."],
    ["previous_marks", "Required. Previous marks percentage, from 0 to 100."],
    ["assignment_score", "Required. Assignment score percentage, from 0 to 100."],
    ["sleep_hours", "Required. Average sleep hours per day, from 0 to 24."],
    [],
    ["Keep the exact column headers in the Students sheet. Add one student per row. Blank rows are ignored."],
    ["Invalid rows are not imported; the app will show the Excel row number and reason."]
  ]);
  instructionsSheet["!cols"] = [{ wch: 24 }, { wch: 82 }];
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, templateSheet, "Students");
  XLSX.utils.book_append_sheet(workbook, instructionsSheet, "Instructions");
  XLSX.writeFile(workbook, "EduPredict_Bulk_Student_Import_Template.xlsx");
}

const navItems = [
  { name: "Dashboard", icon: LayoutDashboard, group: "Workspace" },
  { name: "Students", icon: Users, group: "Workspace" },
  { name: "Student Dashboard", icon: UserRound, group: "Workspace" },
  { name: "AI Prediction", icon: BrainCircuit, group: "Intelligence" },
  { name: "Analytics", icon: BarChart3, group: "Intelligence" },
  { name: "History", icon: Clock3, group: "Management" }
];

function normaliseHistory(item) {
  return {
    ...item,
    student_performance: item.student_performance || item.performance || "Unknown",
    name: item.name || item.student_name || "Student",
    confidence: Number(item.confidence ?? 0),
    date: item.date || item.created_at || ""
  };
}

function formatDate(value) {
  if (!value) return "—";
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? String(value) : date.toLocaleString([], {
    day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit"
  });
}

function downloadCSV(filename, rows) {
  if (!rows.length) return;
  const headers = Object.keys(rows[0]);
  const csv = [headers, ...rows.map(row => headers.map(key => row[key] ?? ""))]
    .map(row => row.map(value => `"${String(value).replaceAll('"', '""')}"`).join(","))
    .join("\n");
  const blob = new Blob(["\ufeff" + csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function App() {
  const [page, setPage] = useState("Dashboard");
  const [students, setStudents] = useState([]);
  const [history, setHistory] = useState([]);
  const [selectedStudentId, setSelectedStudentId] = useState("");
  const [search, setSearch] = useState("");
  const [courseFilter, setCourseFilter] = useState("All courses");
  const [sortBy, setSortBy] = useState("recent");
  const [menuOpen, setMenuOpen] = useState(false);
  const [showForm, setShowForm] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [form, setForm] = useState(initialStudent);
  const [prediction, setPrediction] = useState(null);
  const [loading, setLoading] = useState(false);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const [toast, setToast] = useState("");
  const [darkMode, setDarkMode] = useState(() => localStorage.getItem("edupredict-theme") === "dark");
  const [backendOnline, setBackendOnline] = useState(false);
  const [admin, setAdmin] = useState(null);
  const [authMode, setAuthMode] = useState("login");
  const [authChecking, setAuthChecking] = useState(true);
  const [authLoading, setAuthLoading] = useState(false);
  const [authError, setAuthError] = useState("");
  const [showPassword, setShowPassword] = useState(false);
  const [authForm, setAuthForm] = useState({ name: "", email: "", password: "", confirmPassword: "" });
  const [bulkImportOpen, setBulkImportOpen] = useState(true);
  const [bulkFile, setBulkFile] = useState(null);
  const [bulkDragging, setBulkDragging] = useState(false);
  const [bulkLoading, setBulkLoading] = useState(false);
  const [bulkProgress, setBulkProgress] = useState({ done: 0, total: 0 });
  const [bulkSummary, setBulkSummary] = useState(null);

  const notify = message => setToast(message);

  const loadStudents = async () => {
    try {
      const response = await axios.get(STUDENT_API);
      setStudents(Array.isArray(response.data) ? response.data : []);
      setBackendOnline(true);
    } catch (err) {
      console.error("Student loading failed:", err);
      setBackendOnline(false);
      setError("Unable to load students. Check that the FastAPI backend is running.");
    }
  };

  const loadHistory = async () => {
    try {
      const response = await axios.get(HISTORY_API);
      setHistory(Array.isArray(response.data) ? response.data.map(normaliseHistory) : []);
    } catch (err) {
      console.error("History loading failed:", err);
    }
  };

  useEffect(() => {
    const token = localStorage.getItem(AUTH_TOKEN_KEY);
    if (!token) {
      setAuthChecking(false);
      return;
    }
    axios.get(`${AUTH_API}/me`)
      .then(response => {
        setAdmin(response.data);
        return Promise.all([loadStudents(), loadHistory()]);
      })
      .catch(error => {
        console.error("Admin session validation failed:", error);
        localStorage.removeItem(AUTH_TOKEN_KEY);
        setAdmin(null);
      })
      .finally(() => setAuthChecking(false));
  }, []);

  const handleAuthSubmit = async event => {
    event.preventDefault();
    setAuthError("");
    const email = authForm.email.trim().toLowerCase();
    if (!email || !authForm.password) {
      setAuthError("Please enter your email and password.");
      return;
    }
    if (authMode === "register") {
      if (!authForm.name.trim()) {
        setAuthError("Please enter your full name.");
        return;
      }
      if (authForm.password.length < 8) {
        setAuthError("Use a password with at least 8 characters.");
        return;
      }
      if (authForm.password !== authForm.confirmPassword) {
        setAuthError("The passwords do not match.");
        return;
      }
    }

    setAuthLoading(true);
    try {
      const endpoint = authMode === "register" ? "register" : "login";
      const payload = authMode === "register"
        ? { name: authForm.name.trim(), email, password: authForm.password }
        : { email, password: authForm.password };
      const response = await axios.post(`${AUTH_API}/${endpoint}`, payload);
      localStorage.setItem(AUTH_TOKEN_KEY, response.data.access_token);
      setAdmin(response.data.admin);
      setAuthForm({ name: "", email, password: "", confirmPassword: "" });
      await Promise.all([loadStudents(), loadHistory()]);
      notify(`Welcome to EduPredict, ${response.data.admin.name}.`);
    } catch (error) {
      const detail = error.response?.data?.detail;
      setAuthError(typeof detail === "string" ? detail : "Unable to complete sign in. Check the backend and try again.");
    } finally {
      setAuthLoading(false);
      setAuthChecking(false);
    }
  };

  const handleLogout = () => {
    localStorage.removeItem(AUTH_TOKEN_KEY);
    setAdmin(null);
    setStudents([]);
    setHistory([]);
    setPrediction(null);
    setPage("Dashboard");
    setMenuOpen(false);
    setAuthMode("login");
    setAuthForm({ name: "", email: "", password: "", confirmPassword: "" });
    notify("You have been signed out.");
  };

  useEffect(() => {
    localStorage.setItem("edupredict-theme", darkMode ? "dark" : "light");
  }, [darkMode]);

  useEffect(() => {
    if (!toast) return undefined;
    const timer = window.setTimeout(() => setToast(""), 2800);
    return () => window.clearTimeout(timer);
  }, [toast]);

  const updateForm = event => {
    const { name, value } = event.target;
    setForm(current => ({ ...current, [name]: value }));
  };

  const resetStudentForm = () => {
    setForm(initialStudent);
    setEditingId(null);
    setShowForm(false);
  };

  const saveStudent = async event => {
    event.preventDefault();
    if (!form.name.trim()) {
      notify("Please enter the student's name.");
      return;
    }

    const payload = {
      ...form,
      name: form.name.trim(),
      attendance: Number(form.attendance),
      study_hours: Number(form.study_hours),
      previous_marks: Number(form.previous_marks),
      assignment_score: Number(form.assignment_score),
      sleep_hours: Number(form.sleep_hours)
    };

    if (payload.attendance < 0 || payload.attendance > 100 || payload.previous_marks < 0 || payload.previous_marks > 100 || payload.assignment_score < 0 || payload.assignment_score > 100 || payload.study_hours < 0 || payload.study_hours > 24 || payload.sleep_hours < 0 || payload.sleep_hours > 24) {
      notify("Please check the entered values and their allowed ranges.");
      return;
    }

    setSaving(true);
    try {
      if (editingId !== null) {
        await axios.put(`${STUDENT_API}${editingId}`, payload);
        notify("Student record updated successfully.");
      } else {
        await axios.post(STUDENT_API, payload);
        notify("Student added successfully.");
      }
      await loadStudents();
      resetStudentForm();
    } catch (err) {
      console.error("Saving student failed:", err);
      notify(err.response?.data?.detail || "Could not save student. Check the backend API.");
    } finally {
      setSaving(false);
    }
  };

  const editStudent = student => {
    setForm({ ...initialStudent, ...student });
    setEditingId(student.id);
    setShowForm(true);
    setPage("Students");
  };

  const deleteStudent = async id => {
    if (!window.confirm("Delete this student record?")) return;
    try {
      await axios.delete(`${STUDENT_API}${id}`);
      await loadStudents();
      notify("Student record deleted.");
    } catch (err) {
      console.error("Deleting student failed:", err);
      notify(err.response?.data?.detail || "Could not delete student.");
    }
  };

  const chooseBulkFile = file => {
    if (!file) return;
    const extension = file.name.split(".").pop()?.toLowerCase();
    if (!["xlsx", "xls", "csv"].includes(extension)) {
      notify("Please choose an Excel (.xlsx/.xls) or CSV file.");
      return;
    }
    if (file.size > 10 * 1024 * 1024) {
      notify("File is larger than 10 MB. Please split it into smaller files.");
      return;
    }
    setBulkFile(file);
    setBulkSummary(null);
  };

  const importStudentsFromExcel = async () => {
    if (!bulkFile) {
      notify("Choose an Excel file first.");
      return;
    }

    setBulkLoading(true);
    setBulkSummary(null);
    setBulkProgress({ done: 0, total: 0 });

    try {
      const bytes = await bulkFile.arrayBuffer();
      const workbook = XLSX.read(bytes, { type: "array" });
      const selectedSheetName = workbook.SheetNames.includes("Students")
        ? "Students"
        : workbook.SheetNames[0];
      const worksheet = workbook.Sheets[selectedSheetName];
      if (!worksheet) throw new Error("No readable worksheet was found in this file.");

      const matrix = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        defval: "",
        blankrows: true,
        raw: false
      });
      if (matrix.length < 2) throw new Error("The selected sheet has headers but no student rows.");

      const headers = matrix[0].map(normaliseExcelHeader);
      const columnIndexes = {};
      Object.entries(BULK_COLUMN_ALIASES).forEach(([field, aliases]) => {
        columnIndexes[field] = headers.findIndex(header => aliases.includes(header));
      });

      const missingColumns = REQUIRED_BULK_COLUMNS.filter(field => columnIndexes[field] < 0);
      if (missingColumns.length) {
        throw new Error(`Missing required columns: ${missingColumns.join(", ")}. Download and use the Excel template.`);
      }

      const existingEmails = new Set(
        students.map(student => String(student.email || "").trim().toLowerCase()).filter(Boolean)
      );
      const seenEmails = new Set(existingEmails);
      const validRecords = [];
      const rowErrors = [];
      let skippedDuplicates = 0;
      const dataRows = matrix.slice(1);

      dataRows.forEach((row, index) => {
        const excelRow = index + 2;
        if (!row.some(value => String(value ?? "").trim() !== "")) return;

        const cell = field => {
          const columnIndex = columnIndexes[field];
          return columnIndex >= 0 ? String(row[columnIndex] ?? "").trim() : "";
        };

        const name = cell("name");
        const email = cell("email").toLowerCase();
        const course = cell("course");
        if (!name) {
          rowErrors.push(`Row ${excelRow}: Student name is required.`);
          return;
        }
        if (!course) {
          rowErrors.push(`Row ${excelRow}: Course is required.`);
          return;
        }
        if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
          rowErrors.push(`Row ${excelRow}: Email address is invalid.`);
          return;
        }
        if (email && seenEmails.has(email)) {
          skippedDuplicates += 1;
          return;
        }

        const numericRules = [
          ["attendance", 0, 100],
          ["study_hours", 0, 24],
          ["previous_marks", 0, 100],
          ["assignment_score", 0, 100],
          ["sleep_hours", 0, 24]
        ];
        const record = { name, email, course };
        let invalidReason = "";
        for (const [field, min, max] of numericRules) {
          const rawValue = cell(field);
          const number = Number(rawValue);
          if (rawValue === "" || !Number.isFinite(number)) {
            invalidReason = `${field} must be a valid number.`;
            break;
          }
          if (number < min || number > max) {
            invalidReason = `${field} must be between ${min} and ${max}.`;
            break;
          }
          record[field] = number;
        }

        if (invalidReason) {
          rowErrors.push(`Row ${excelRow}: ${invalidReason}`);
          return;
        }

        validRecords.push({ row: excelRow, record });
        if (email) seenEmails.add(email);
      });

      if (validRecords.length > 500) {
        throw new Error("A maximum of 500 students can be imported at once. Split the workbook into smaller files.");
      }

      setBulkProgress({ done: 0, total: validRecords.length });
      const failedRows = [];
      let added = 0;

      for (let index = 0; index < validRecords.length; index += 1) {
        const item = validRecords[index];
        try {
          await axios.post(STUDENT_API, item.record);
          added += 1;
        } catch (requestError) {
          const detail = requestError.response?.data?.detail;
          failedRows.push(`Row ${item.row}: ${typeof detail === "string" ? detail : "Could not save this student record."}`);
        }
        setBulkProgress({ done: index + 1, total: validRecords.length });
      }

      await loadStudents();
      const errors = [...rowErrors, ...failedRows];
      setBulkSummary({
        added,
        skipped: skippedDuplicates,
        invalid: rowErrors.length,
        failed: failedRows.length,
        total: dataRows.filter(row => row.some(value => String(value ?? "").trim() !== "")).length,
        errors: errors.slice(0, 30),
        hasMoreErrors: errors.length > 30
      });
      notify(`Excel import completed: ${added} student(s) added.`);
    } catch (importError) {
      console.error("Bulk student import failed:", importError);
      notify(importError.message || "Unable to read this Excel file.");
    } finally {
      setBulkLoading(false);
    }
  };

  const openPredictionFor = student => {
    setForm({ ...initialStudent, ...student });
    setPrediction(null);
    setError("");
    setPage("AI Prediction");
  };

  const predict = async event => {
    event.preventDefault();
    setLoading(true);
    setError("");
    setPrediction(null);
    const payload = {
      name: form.name.trim() || "Student",
      attendance: Number(form.attendance),
      study_hours: Number(form.study_hours),
      previous_marks: Number(form.previous_marks),
      assignment_score: Number(form.assignment_score),
      sleep_hours: Number(form.sleep_hours)
    };

    try {
      const response = await axios.post(API, payload);
      const result = {
        ...response.data,
        student_performance: response.data.student_performance || response.data.performance,
        confidence: Number(response.data.confidence ?? 0),
        name: payload.name,
        date: new Date().toISOString()
      };
      setPrediction(result);
      await loadHistory();
      notify("AI prediction generated successfully.");
    } catch (err) {
      console.error("Prediction failed:", err);
      setError(err.response?.data?.detail ? JSON.stringify(err.response.data.detail) : "Backend connection failed. Make sure FastAPI is running on port 8000.");
    } finally {
      setLoading(false);
    }
  };

  const totalStudents = students.length;
  const averageAttendance = totalStudents ? (students.reduce((sum, student) => sum + Number(student.attendance || 0), 0) / totalStudents) : 0;
  const averageMarks = totalStudents ? (students.reduce((sum, student) => sum + Number(student.previous_marks || 0), 0) / totalStudents) : 0;
  const averageStudyHours = totalStudents ? (students.reduce((sum, student) => sum + Number(student.study_hours || 0), 0) / totalStudents) : 0;
  const excellent = history.filter(item => item.student_performance === "Excellent").length;
  const average = history.filter(item => item.student_performance === "Average").length;
  const atRisk = history.filter(item => item.student_performance === "At Risk").length;
  const selectedStudent = students.find(student => String(student.id) === String(selectedStudentId)) || students[0] || null;
  const selectedStudentHistory = selectedStudent
    ? history.filter(item => String(item.name || "").trim().toLowerCase() === String(selectedStudent.name || "").trim().toLowerCase())
      .sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime())
    : [];
  const latestStudentPrediction = selectedStudentHistory[0] || null;
  const courseOptions = useMemo(() => ["All courses", ...new Set(students.map(student => student.course).filter(Boolean))], [students]);

  const filteredStudents = useMemo(() => {
    const query = search.trim().toLowerCase();
    const filtered = students.filter(student => {
      const matchesSearch = !query || [student.name, student.email, student.course].some(value => String(value || "").toLowerCase().includes(query));
      const matchesCourse = courseFilter === "All courses" || student.course === courseFilter;
      return matchesSearch && matchesCourse;
    });
    return [...filtered].sort((a, b) => {
      if (sortBy === "name") return String(a.name || "").localeCompare(String(b.name || ""));
      if (sortBy === "marks") return Number(b.previous_marks || 0) - Number(a.previous_marks || 0);
      if (sortBy === "attendance") return Number(b.attendance || 0) - Number(a.attendance || 0);
      return Number(b.id || 0) - Number(a.id || 0);
    });
  }, [students, search, courseFilter, sortBy]);

  const distribution = [
    { name: "Excellent", value: excellent, color: performanceColors.Excellent },
    { name: "Average", value: average, color: performanceColors.Average },
    { name: "At Risk", value: atRisk, color: performanceColors["At Risk"] }
  ];

  const marksBands = [
    { band: "Below 50", count: students.filter(s => Number(s.previous_marks) < 50).length },
    { band: "50–69", count: students.filter(s => Number(s.previous_marks) >= 50 && Number(s.previous_marks) < 70).length },
    { band: "70–84", count: students.filter(s => Number(s.previous_marks) >= 70 && Number(s.previous_marks) < 85).length },
    { band: "85+", count: students.filter(s => Number(s.previous_marks) >= 85).length }
  ];
  const attendanceBands = [
    { band: "Below 50", count: students.filter(s => Number(s.attendance) < 50).length },
    { band: "50–64", count: students.filter(s => Number(s.attendance) >= 50 && Number(s.attendance) < 65).length },
    { band: "65–74", count: students.filter(s => Number(s.attendance) >= 65 && Number(s.attendance) < 75).length },
    { band: "75–89", count: students.filter(s => Number(s.attendance) >= 75 && Number(s.attendance) < 90).length },
    { band: "90+", count: students.filter(s => Number(s.attendance) >= 90).length }
  ];
  const studyHourBands = [
    { band: "0–1.9 h", count: students.filter(s => Number(s.study_hours) < 2).length },
    { band: "2–3.9 h", count: students.filter(s => Number(s.study_hours) >= 2 && Number(s.study_hours) < 4).length },
    { band: "4–5.9 h", count: students.filter(s => Number(s.study_hours) >= 4 && Number(s.study_hours) < 6).length },
    { band: "6–7.9 h", count: students.filter(s => Number(s.study_hours) >= 6 && Number(s.study_hours) < 8).length },
    { band: "8+ h", count: students.filter(s => Number(s.study_hours) >= 8).length }
  ];

  const recentHistory = [...history].sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime()).slice(0, 5);
  const attentionStudents = students.filter(student => Number(student.attendance) < 65 || Number(student.previous_marks) < 50).slice(0, 4);

  const exportStudents = () => downloadCSV("edupredict-students.csv", students.map(({ id, name, email, course, attendance, study_hours, previous_marks, assignment_score, sleep_hours }) => ({ id, name, email, course, attendance, study_hours, previous_marks, assignment_score, sleep_hours })));
  const exportHistory = () => downloadCSV("edupredict-prediction-history.csv", history.map(item => ({ name: item.name, performance: item.student_performance, confidence: item.confidence, date: item.date })));

  const reportRows = students.map(student => {
    const matchingHistory = history.filter(item => String(item.name || "").trim().toLowerCase() === String(student.name || "").trim().toLowerCase())
      .sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime());
    const latest = matchingHistory[0] || null;
    return {
      Student_ID: student.id,
      Student_Name: student.name,
      Email: student.email || "",
      Course: student.course || "",
      Attendance_Percent: Number(student.attendance || 0),
      Study_Hours_Per_Day: Number(student.study_hours || 0),
      Previous_Marks_Percent: Number(student.previous_marks || 0),
      Assignment_Score_Percent: Number(student.assignment_score || 0),
      Sleep_Hours_Per_Day: Number(student.sleep_hours || 0),
      Latest_Performance_Category: latest?.student_performance || "Not predicted",
      Latest_Model_Confidence_Percent: latest ? Number(latest.confidence || 0) : "",
      Latest_Prediction_Date: latest?.date || "",
      Prediction_Count: matchingHistory.length
    };
  });

  const exportPerformanceReport = () => {
    if (!reportRows.length) {
      notify("Add student records before downloading the performance report.");
      return;
    }
    downloadCSV("edupredict-performance-report.csv", reportRows);
    notify("Performance report downloaded as CSV.");
  };

  const printPerformanceReport = () => {
    if (!reportRows.length) {
      notify("Add student records before creating a report.");
      return;
    }
    const escapeHTML = value => String(value ?? "").replace(/[&<>"']/g, char => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#39;" }[char]));
    const rowsHTML = reportRows.map(row => `<tr><td>${escapeHTML(row.Student_Name)}</td><td>${escapeHTML(row.Course)}</td><td>${row.Attendance_Percent}%</td><td>${row.Study_Hours_Per_Day} h</td><td>${row.Previous_Marks_Percent}%</td><td>${escapeHTML(row.Latest_Performance_Category)}</td><td>${row.Latest_Model_Confidence_Percent === "" ? "—" : `${row.Latest_Model_Confidence_Percent}%`}</td></tr>`).join("");
    const reportWindow = window.open("", "_blank", "width=1100,height=800");
    if (!reportWindow) {
      notify("Allow pop-ups to print or save the report as PDF.");
      return;
    }
    reportWindow.document.write(`<!doctype html><html><head><title>EduPredict Performance Report</title><meta charset="utf-8"><style>body{font-family:Arial,sans-serif;color:#20243a;padding:32px}h1{margin:0 0 6px;color:#4f4bb5}p{color:#68708a;font-size:12px}.summary{display:flex;gap:12px;margin:22px 0}.summary div{border:1px solid #e5e7ef;padding:12px 16px;border-radius:9px;min-width:120px}.summary b{display:block;font-size:20px;margin-top:5px}table{width:100%;border-collapse:collapse;margin-top:20px;font-size:11px}th,td{border:1px solid #e5e7ef;padding:9px;text-align:left}th{background:#f1f1ff;color:#4b4b80}.footer{margin-top:25px;color:#8890a5;font-size:10px}@media print{body{padding:0}button{display:none}}</style></head><body><h1>EduPredict</h1><p>Student Performance Report · Generated ${escapeHTML(new Date().toLocaleString())}</p><p>Prepared for: ${escapeHTML(admin?.name || "Administrator")} (${escapeHTML(admin?.email || "")})</p><div class="summary"><div>Total Students<b>${students.length}</b></div><div>Avg Attendance<b>${averageAttendance.toFixed(1)}%</b></div><div>Avg Study Hours<b>${averageStudyHours.toFixed(1)} h</b></div><div>Avg Previous Marks<b>${averageMarks.toFixed(1)}%</b></div><div>AI Predictions<b>${history.length}</b></div></div><table><thead><tr><th>Student</th><th>Course</th><th>Attendance</th><th>Study / day</th><th>Previous marks</th><th>Latest category</th><th>Confidence</th></tr></thead><tbody>${rowsHTML}</tbody></table><p class="footer">Performance categories are produced by the configured machine-learning model. Confidence is the model-reported score and is not a guarantee of future academic outcome.</p><button onclick="window.print()">Print / Save as PDF</button><script>window.onload=()=>window.print()</script></body></html>`);
    reportWindow.document.close();
  };

  const goTo = name => {
    setPage(name);
    setMenuOpen(false);
    setError("");
  };

  const openStudentDashboard = student => {
    if (!student) return;
    setSelectedStudentId(String(student.id));
    goTo("Student Dashboard");
  };

  if (authChecking) {
    return <div className="auth-loading-screen"><div className="auth-loading-mark"><GraduationCap size={28}/></div><div className="auth-loading-spinner"/><p>Securing your admin workspace…</p></div>;
  }

  if (!admin) {
    return <AuthScreen
      mode={authMode}
      setMode={mode => { setAuthMode(mode); setAuthError(""); }}
      form={authForm}
      setForm={setAuthForm}
      onSubmit={handleAuthSubmit}
      loading={authLoading}
      error={authError}
      showPassword={showPassword}
      setShowPassword={setShowPassword}
    />;
  }

  return (
    <div className={`app-shell ${darkMode ? "theme-dark" : "theme-light"}`}>
      <aside className={`sidebar ${menuOpen ? "sidebar-open" : ""}`}>
        <div className="brand-lockup">
          <div className="brand-mark"><GraduationCap size={23} strokeWidth={2.2}/></div>
          <div className="brand-copy"><strong>EduPredict</strong><span>STUDENT INTELLIGENCE</span></div>
          <button className="sidebar-close icon-btn" onClick={() => setMenuOpen(false)} aria-label="Close menu"><X size={18}/></button>
        </div>

        <div className="workspace-switcher">
          <div className="workspace-avatar">ED</div>
          <div><b>{admin?.name || "Academic Workspace"}</b><span>{admin?.email || "Administrator account"}</span></div>
          <ChevronDown size={15}/>
        </div>

        {['Workspace', 'Intelligence', 'Management'].map(group => (
          <div className="nav-group" key={group}>
            <p className="nav-caption">{group}</p>
            {navItems.filter(item => item.group === group).map(item => {
              const Icon = item.icon;
              return (
                <button key={item.name} className={`nav-link ${page === item.name ? "active" : ""}`} onClick={() => goTo(item.name)}>
                  <Icon size={18}/><span>{item.name}</span>{page === item.name && <span className="active-indicator"/>}
                  {item.name === "History" && history.length > 0 && <small>{history.length}</small>}
                </button>
              );
            })}
          </div>
        ))}

        <div className="sidebar-bottom">
          <div className="sidebar-insight">
            <div className="insight-orb"><Sparkles size={17}/></div>
            <b>Make data-driven decisions</b>
            <p>Use attendance and academic patterns to identify students who may need support.</p>
            <button onClick={() => goTo("Analytics")}>Explore insights <ArrowUpRight size={14}/></button>
          </div>
          <div className="sidebar-user">
            <div className="user-avatar">{getInitials(admin?.name)}</div>
            <div><b>{admin?.name || "Administrator"}</b><span>System Administrator</span></div>
            <Settings2 size={17}/>
          </div>
        </div>
      </aside>

      {menuOpen && <button className="mobile-backdrop" onClick={() => setMenuOpen(false)} aria-label="Close navigation"/>}

      <main className="main-area">
        <header className="topbar">
          <div className="topbar-left">
            <button className="icon-btn mobile-menu" onClick={() => setMenuOpen(true)} aria-label="Open menu"><Menu size={20}/></button>
            <div className="breadcrumb"><span>EduPredict</span><b>/</b><strong>{page}</strong></div>
          </div>
          <div className="topbar-right">
            <div className="date-chip"><CalendarDays size={15}/>{new Date().toLocaleDateString([], { day: "2-digit", month: "short", year: "numeric" })}</div>
            <div className={`online-pill ${backendOnline ? "" : "status-offline"}`}><span/> {backendOnline ? "Backend connected" : "Backend offline"}</div>
            <button className="icon-btn" onClick={() => setDarkMode(value => !value)} title={darkMode ? "Switch to light mode" : "Switch to dark mode"}>{darkMode ? <Sun size={18}/> : <Moon size={18}/>}</button>
            <div className="top-avatar" title={admin?.email || "Administrator"}>{getInitials(admin?.name)}</div>
            <button className="icon-btn logout-btn" onClick={handleLogout} title="Sign out" aria-label="Sign out"><LogOut size={17}/></button>
          </div>
        </header>

        <div className="page-content">
          {page === "Dashboard" && (
            <>
              <section className="welcome-banner">
                <div className="welcome-copy">
                  <span className="section-kicker"><span className="kicker-dot"/> STUDENT ANALYTICS PLATFORM</span>
                  <h1>Good to see you, {admin?.name?.split(" ")[0] || "Admin"} <span>✦</span></h1>
                  <p>Your academic overview, predictions and student insights — all in one place.</p>
                  <div className="welcome-actions">
                    <button className="btn-primary" onClick={() => { resetStudentForm(); setPage("Students"); setShowForm(true); }}><Plus size={17}/> Add student</button>
                    <button className="btn-ghost" onClick={() => goTo("AI Prediction")}><BrainCircuit size={17}/> Run AI prediction <ArrowUpRight size={15}/></button>
                    <button className="btn-ghost" onClick={exportPerformanceReport} disabled={!students.length}><Download size={16}/> Performance report</button>
                  </div>
                </div>
                <div className="welcome-art" aria-hidden="true">
                  <div className="art-ring ring-one"/><div className="art-ring ring-two"/>
                  <div className="art-core"><BrainCircuit size={49}/></div>
                  <div className="art-tag tag-one"><TrendingUp size={14}/> Smart insights</div>
                  <div className="art-tag tag-two"><ShieldCheck size={14}/> AI powered</div>
                </div>
              </section>

              <section className="kpi-grid">
                <KpiCard icon={Users} label="Total students" value={totalStudents} note="Registered records" tone="violet"/>
                <KpiCard icon={UserRound} label="Average attendance" value={totalStudents ? `${averageAttendance.toFixed(1)}%` : "—"} note="Across all students" tone="mint"/>
                <KpiCard icon={Award} label="Average marks" value={totalStudents ? `${averageMarks.toFixed(1)}%` : "—"} note="Previous academic marks" tone="amber"/>
                <KpiCard icon={BrainCircuit} label="AI predictions" value={history.length} note={`${atRisk} marked at risk`} tone="blue"/>
              </section>

              <section className="dashboard-grid">
                <div className="surface chart-surface">
                  <div className="surface-heading"><div><span className="mini-label">MODEL OUTPUT</span><h2>Performance distribution</h2><p>Classification from saved AI predictions</p></div><span className="chart-icon"><PieChartIcon/></span></div>
                  {history.length ? (
                    <div className="distribution-layout">
                      <div className="donut-wrap"><ResponsiveContainer width="100%" height={230}><PieChart><Pie data={distribution} dataKey="value" nameKey="name" innerRadius={66} outerRadius={91} paddingAngle={4} stroke="none">{distribution.map(entry => <Cell key={entry.name} fill={entry.color}/>)}</Pie><Tooltip/></PieChart></ResponsiveContainer><div className="donut-center"><strong>{history.length}</strong><span>Predictions</span></div></div>
                      <div className="legend-list">{distribution.map(item => <div className="legend-row" key={item.name}><span className="legend-dot" style={{ background: item.color }}/><span>{item.name}</span><b>{item.value}</b><small>{history.length ? `${Math.round(item.value / history.length * 100)}%` : "0%"}</small></div>)}</div>
                    </div>
                  ) : <EmptyState icon={BrainCircuit} title="No predictions yet" text="Run your first prediction to see the model output distribution." action={() => goTo("AI Prediction")} actionText="Start prediction"/>}
                </div>

                <div className="surface chart-surface">
                  <div className="surface-heading"><div><span className="mini-label">ACADEMIC SNAPSHOT</span><h2>Marks overview</h2><p>Student count by previous marks range</p></div><span className="chart-icon"><BarChart3 size={18}/></span></div>
                  {students.length ? <div className="wide-chart"><ResponsiveContainer width="100%" height={250}><BarChart data={marksBands} margin={{ top: 12, right: 8, left: -18, bottom: 0 }}><CartesianGrid strokeDasharray="3 3" vertical={false} stroke="var(--chart-grid)"/><XAxis dataKey="band" axisLine={false} tickLine={false} tick={{ fill: "var(--muted)", fontSize: 11 }}/><YAxis allowDecimals={false} axisLine={false} tickLine={false} tick={{ fill: "var(--muted)", fontSize: 11 }}/><Tooltip cursor={{ fill: "var(--hover)" }}/><Bar dataKey="count" name="Students" fill="#7775ef" radius={[7, 7, 0, 0]} barSize={34}/></BarChart></ResponsiveContainer></div> : <EmptyState icon={BarChart3} title="No student data" text="Add student records to populate this chart." action={() => { setPage("Students"); setShowForm(true); }} actionText="Add student"/>}
                </div>
              </section>

              <section className="dashboard-grid lower-grid">
                <div className="surface table-surface">
                  <div className="surface-heading"><div><span className="mini-label">STUDENT DIRECTORY</span><h2>Recently added students</h2><p>Your latest registered student records</p></div><button className="text-action" onClick={() => goTo("Students")}>View all <ArrowUpRight size={15}/></button></div>
                  <StudentTable students={[...students].sort((a, b) => Number(b.id || 0) - Number(a.id || 0)).slice(0, 5)} onEdit={editStudent} onDelete={deleteStudent} onAnalyze={openPredictionFor} onView={openStudentDashboard}/>
                </div>
                <div className="surface activity-surface">
                  <div className="surface-heading"><div><span className="mini-label">RECENT ACTIVITY</span><h2>Latest predictions</h2><p>Most recent AI analysis results</p></div><button className="icon-btn subtle" onClick={() => goTo("History")} title="View history"><ArrowUpRight size={17}/></button></div>
                  {recentHistory.length ? <div className="activity-list">{recentHistory.map((item, index) => <div className="activity-item" key={`${item.name}-${item.date}-${index}`}><div className={`activity-icon ${performanceClass(item.student_performance)}`}><BrainCircuit size={16}/></div><div className="activity-copy"><b>{item.name}</b><span>Predicted as <strong>{item.student_performance}</strong></span><small>{formatDate(item.date)}</small></div><span className={`confidence-mini ${performanceClass(item.student_performance)}`}>{Number(item.confidence).toFixed(0)}%</span></div>)}</div> : <EmptyState icon={Activity} title="Activity will appear here" text="Your recent prediction records will show up here."/>}
                </div>
              </section>

              {attentionStudents.length > 0 && <section className="attention-strip"><div className="attention-symbol"><AlertCircle size={19}/></div><div><b>{attentionStudents.length} student{attentionStudents.length > 1 ? "s" : ""} may need attention</b><span>Based on attendance below 65% or previous marks below 50%. This is a simple screening rule, not an AI prediction.</span></div><button onClick={() => { setCourseFilter("All courses"); setSearch(""); goTo("Students"); }}>Review students <ArrowUpRight size={15}/></button></section>}
            </>
          )}

          {page === "Students" && (
            <>
              <PageTitle eyebrow="STUDENT DIRECTORY" title="Manage students" subtitle="Maintain student profiles and launch predictions directly from a record." actions={<><button className="btn-secondary" onClick={() => { setBulkImportOpen(open => !open); setBulkSummary(null); }}><FileSpreadsheet size={16}/> Import Excel</button><button className="btn-secondary" onClick={exportStudents} disabled={!students.length}><Download size={16}/> Export CSV</button><button className="btn-primary" onClick={() => { setForm(initialStudent); setEditingId(null); setShowForm(true); }}><Plus size={17}/> Add student</button></>}/>
              <div className="directory-summary"><div><Users size={18}/><span><b>{students.length}</b> total records</span></div><div><BookOpen size={18}/><span><b>{courseOptions.length - 1}</b> courses</span></div><div><Target size={18}/><span><b>{students.filter(s => Number(s.attendance) >= 75).length}</b> with 75%+ attendance</span></div></div>
              {bulkImportOpen && <section className="surface bulk-import-surface">
                <div className="surface-heading bulk-import-heading">
                  <div><span className="mini-label">BULK REGISTRATION</span><h2>Import students from Excel</h2><p>Upload a workbook and register many student records in one operation.</p></div>
                  <button className="icon-btn" onClick={() => { setBulkImportOpen(false); setBulkFile(null); setBulkSummary(null); }} title="Close import panel"><X size={18}/></button>
                </div>
                <div
                  className={`bulk-dropzone ${bulkDragging ? "is-dragging" : ""} ${bulkFile ? "has-file" : ""}`}
                  onDragOver={event => { event.preventDefault(); setBulkDragging(true); }}
                  onDragLeave={event => { event.preventDefault(); setBulkDragging(false); }}
                  onDrop={event => { event.preventDefault(); setBulkDragging(false); chooseBulkFile(event.dataTransfer.files?.[0]); }}
                >
                  <input className="bulk-file-input" id="bulk-students-file" type="file" accept=".xlsx,.xls,.csv" onChange={event => chooseBulkFile(event.target.files?.[0])}/>
                  <div className="bulk-drop-icon"><FileSpreadsheet size={27}/></div>
                  <div className="bulk-drop-copy"><b>{bulkFile ? bulkFile.name : "Drag & drop your Excel file here"}</b><span>{bulkFile ? `${(bulkFile.size / 1024).toFixed(1)} KB · Ready to import` : "or choose a file from your computer · XLSX, XLS or CSV · Max 10 MB"}</span></div>
                  <label className="btn-secondary bulk-browse-button" htmlFor="bulk-students-file">Browse files</label>
                </div>
                <div className="bulk-import-meta">
                  <div><b>Required fields</b><span>name, course, attendance, study_hours, previous_marks, assignment_score, sleep_hours</span></div>
                  <div><b>Optional field</b><span>email (used to identify duplicate student records)</span></div>
                  <button className="text-action" type="button" onClick={downloadExcelTemplate}><Download size={15}/> Download Excel template</button>
                </div>
                {bulkLoading && <div className="bulk-progress-wrap"><div className="bulk-progress-label"><span>Importing students…</span><b>{bulkProgress.done} / {bulkProgress.total}</b></div><div className="bulk-progress-track"><span style={{ width: `${bulkProgress.total ? (bulkProgress.done / bulkProgress.total) * 100 : 0}%` }}/></div></div>}
                <div className="bulk-import-actions"><span>Duplicate emails are skipped. Invalid rows are reported and not imported.</span><button className="btn-primary" type="button" onClick={importStudentsFromExcel} disabled={!bulkFile || bulkLoading}>{bulkLoading ? <><span className="button-spinner"/> Importing…</> : <><FileSpreadsheet size={16}/> Import students</>}</button></div>
                {bulkSummary && <div className="bulk-summary-panel">
                  <div className="bulk-summary-title"><Check size={17}/><b>Import summary</b><span>{bulkSummary.total} data rows checked</span></div>
                  <div className="bulk-summary-stats"><div><b>{bulkSummary.added}</b><span>Added</span></div><div><b>{bulkSummary.skipped}</b><span>Duplicates skipped</span></div><div><b>{bulkSummary.invalid}</b><span>Invalid rows</span></div><div><b>{bulkSummary.failed}</b><span>Save failures</span></div></div>
                  {bulkSummary.errors.length > 0 && <div className="bulk-error-list"><b>Rows needing attention</b>{bulkSummary.errors.map((message, index) => <span key={`${index}-${message}`}>{message}</span>)}{bulkSummary.hasMoreErrors && <small>Only the first 30 errors are shown.</small>}</div>}
                </div>}
              </section>}

              {showForm && <div className="modal-backdrop" onMouseDown={event => { if (event.target === event.currentTarget) resetStudentForm(); }}><form className="student-modal" onSubmit={saveStudent}><div className="modal-heading"><div><span className="mini-label">STUDENT PROFILE</span><h2>{editingId !== null ? "Edit student" : "Add new student"}</h2><p>Enter academic and contact details below.</p></div><button type="button" className="icon-btn" onClick={resetStudentForm}><X size={19}/></button></div><div className="form-grid"><Field label="Full name" name="name" value={form.name} onChange={updateForm} placeholder="e.g. Rahul Sharma"/><Field label="Email address" name="email" value={form.email} onChange={updateForm} type="email" placeholder="student@example.com" required={false}/><Field label="Course" name="course" value={form.course} onChange={updateForm} placeholder="e.g. BCA AI & DS"/><Field label="Attendance (%)" name="attendance" value={form.attendance} onChange={updateForm} type="number" min="0" max="100"/><Field label="Study hours / day" name="study_hours" value={form.study_hours} onChange={updateForm} type="number" min="0" max="24" step="0.5"/><Field label="Previous marks (%)" name="previous_marks" value={form.previous_marks} onChange={updateForm} type="number" min="0" max="100"/><Field label="Assignment score (%)" name="assignment_score" value={form.assignment_score} onChange={updateForm} type="number" min="0" max="100"/><Field label="Sleep hours / day" name="sleep_hours" value={form.sleep_hours} onChange={updateForm} type="number" min="0" max="24" step="0.5"/></div><div className="modal-footer"><button type="button" className="btn-secondary" onClick={resetStudentForm}>Cancel</button><button type="submit" className="btn-primary" disabled={saving}>{saving ? "Saving…" : <><Check size={16}/> {editingId !== null ? "Save changes" : "Create student"}</>}</button></div></form></div>}

              <div className="surface directory-surface">
                <div className="directory-toolbar"><div><h2>All students</h2><p>Search, filter and manage student records</p></div><div className="toolbar-controls"><label className="search-control"><Search size={16}/><input value={search} onChange={event => setSearch(event.target.value)} placeholder="Search name, email or course…"/></label><label className="select-control"><Filter size={15}/><select value={courseFilter} onChange={event => setCourseFilter(event.target.value)}>{courseOptions.map(option => <option key={option}>{option}</option>)}</select></label><label className="select-control sort-control"><select value={sortBy} onChange={event => setSortBy(event.target.value)}><option value="recent">Latest added</option><option value="name">Name A–Z</option><option value="marks">Highest marks</option><option value="attendance">Highest attendance</option></select><ChevronDown size={14}/></label></div></div>
                <StudentTable students={filteredStudents} onEdit={editStudent} onDelete={deleteStudent} onAnalyze={openPredictionFor} onView={openStudentDashboard}/>
                <div className="table-footer">Showing <b>{filteredStudents.length}</b> of <b>{students.length}</b> students</div>
              </div>
            </>
          )}

          {page === "Student Dashboard" && (
            <>
              <PageTitle eyebrow="INDIVIDUAL STUDENT OVERVIEW" title="Student dashboard" subtitle="Inspect an individual student's academic indicators and latest prediction." actions={<label className="select-control student-dashboard-picker"><UserRound size={15}/><select value={selectedStudent ? String(selectedStudent.id) : ""} onChange={event => setSelectedStudentId(event.target.value)} disabled={!students.length}><option value="">Select a student</option>{students.map(student => <option key={student.id} value={String(student.id)}>{student.name} · {student.course || "No course"}</option>)}</select></label>}/>
              {selectedStudent ? (
                <>
                  <section className="student-profile-banner">
                    <div className="student-profile-avatar">{getInitials(selectedStudent.name)}</div>
                    <div className="student-profile-main"><span className="mini-label">STUDENT PROFILE · ID #{selectedStudent.id}</span><h2>{selectedStudent.name}</h2><p>{selectedStudent.email || "No email provided"} <span>•</span> {selectedStudent.course || "Course not specified"}</p></div>
                    <div className="student-profile-actions"><button className="btn-secondary" onClick={() => editStudent(selectedStudent)}><Pencil size={15}/> Edit profile</button><button className="btn-primary" onClick={() => openPredictionFor(selectedStudent)}><BrainCircuit size={15}/> Run prediction</button></div>
                  </section>
                  <section className="kpi-grid student-dashboard-kpis">
                    <KpiCard icon={UserRound} label="Attendance" value={`${Number(selectedStudent.attendance || 0).toFixed(1)}%`} note="Recorded attendance" tone="mint"/>
                    <KpiCard icon={Clock3} label="Study hours" value={`${Number(selectedStudent.study_hours || 0).toFixed(1)} h`} note="Per day" tone="blue"/>
                    <KpiCard icon={Award} label="Previous marks" value={`${Number(selectedStudent.previous_marks || 0).toFixed(1)}%`} note="Academic marks" tone="amber"/>
                    <KpiCard icon={BrainCircuit} label="Saved predictions" value={selectedStudentHistory.length} note={latestStudentPrediction ? `Latest: ${latestStudentPrediction.student_performance}` : "No prediction yet"} tone="violet"/>
                  </section>
                  <section className="dashboard-grid student-dashboard-grid">
                    <div className="surface chart-surface">
                      <div className="surface-heading"><div><span className="mini-label">ACADEMIC INDICATORS</span><h2>Attendance & marks profile</h2><p>Current recorded academic percentages</p></div></div>
                      <ResponsiveContainer width="100%" height={275}><BarChart data={[{ metric: "Attendance", score: Number(selectedStudent.attendance || 0) }, { metric: "Previous marks", score: Number(selectedStudent.previous_marks || 0) }, { metric: "Assignment", score: Number(selectedStudent.assignment_score || 0) }]} margin={{ top: 14, right: 12, left: -12, bottom: 5 }}><CartesianGrid strokeDasharray="3 3" vertical={false} stroke="var(--chart-grid)"/><XAxis dataKey="metric" axisLine={false} tickLine={false} tick={{ fill: "var(--muted)", fontSize: 10 }}/><YAxis domain={[0, 100]} axisLine={false} tickLine={false} tick={{ fill: "var(--muted)", fontSize: 10 }}/><Tooltip formatter={value => [`${value}%`, "Score"]}/><Bar dataKey="score" name="Score" fill="#7775ef" radius={[7, 7, 0, 0]} barSize={48}/></BarChart></ResponsiveContainer>
                    </div>
                    <div className="surface chart-surface student-routine-card">
                      <div className="surface-heading"><div><span className="mini-label">DAILY ROUTINE</span><h2>Study & rest snapshot</h2><p>Reported average hours per day</p></div></div>
                      <div className="routine-metric"><div className="routine-icon study"><BookOpen size={18}/></div><div><span>Study hours / day</span><b>{Number(selectedStudent.study_hours || 0).toFixed(1)} h</b></div><small>of 24 h</small></div>
                      <div className="routine-track"><span style={{ width: `${Math.min(100, Number(selectedStudent.study_hours || 0) / 24 * 100)}%` }}/></div>
                      <div className="routine-metric"><div className="routine-icon sleep"><Clock3 size={18}/></div><div><span>Sleep hours / day</span><b>{Number(selectedStudent.sleep_hours || 0).toFixed(1)} h</b></div><small>of 24 h</small></div>
                      <div className="routine-track sleep-track"><span style={{ width: `${Math.min(100, Number(selectedStudent.sleep_hours || 0) / 24 * 100)}%` }}/></div>
                      <div className="routine-metric"><div className="routine-icon assignment"><FileSpreadsheet size={18}/></div><div><span>Assignment score</span><b>{Number(selectedStudent.assignment_score || 0).toFixed(1)}%</b></div></div>
                      {latestStudentPrediction ? <div className="latest-prediction-box"><span>Latest model category</span><PerformanceBadge value={latestStudentPrediction.student_performance}/><b>{Number(latestStudentPrediction.confidence || 0).toFixed(1)}% confidence</b></div> : <div className="no-prediction-box">No prediction has been saved for this student yet.</div>}
                    </div>
                  </section>
                  <section className="surface table-surface student-history-surface"><div className="surface-heading"><div><span className="mini-label">INDIVIDUAL AUDIT TRAIL</span><h2>Prediction history for {selectedStudent.name}</h2><p>Saved model outputs matching this student's name</p></div><button className="text-action" onClick={() => openPredictionFor(selectedStudent)}>New prediction <ArrowUpRight size={14}/></button></div>{selectedStudentHistory.length ? <div className="table-scroll"><table className="data-table"><thead><tr><th>DATE & TIME</th><th>PERFORMANCE CATEGORY</th><th>MODEL CONFIDENCE</th></tr></thead><tbody>{selectedStudentHistory.map((item, index) => <tr key={`${item.date}-${index}`}><td className="date-cell">{formatDate(item.date)}</td><td><PerformanceBadge value={item.student_performance}/></td><td><div className="confidence-cell"><div className="confidence-track"><span style={{ width: `${Math.max(0, Math.min(100, item.confidence))}%` }}/></div><b>{Number(item.confidence || 0).toFixed(1)}%</b></div></td></tr>)}</tbody></table></div> : <EmptyState icon={Clock3} title="No saved predictions" text="Run a prediction for this student to build an individual history." action={() => openPredictionFor(selectedStudent)} actionText="Run prediction"/>}</section>
                </>
              ) : <div className="surface"><EmptyState icon={Users} title="No student registered yet" text="Register a student first. Their individual dashboard will appear here." action={() => goTo("Students")} actionText="Register student"/></div>}
            </>
          )}

          {page === "AI Prediction" && (
            <>
              <PageTitle eyebrow="MACHINE LEARNING WORKSPACE" title="AI performance prediction" subtitle="Evaluate a student's academic indicators using your trained Random Forest model."/>
              <div className="prediction-layout">
                <section className="surface prediction-form-surface"><div className="surface-heading"><div><span className="mini-label">INPUT PARAMETERS</span><h2>Student performance profile</h2><p>Complete all academic indicators to generate a prediction.</p></div><div className="model-badge"><span/> RANDOM FOREST</div></div>
                  <form onSubmit={predict}>
                    <div className="form-grid"><Field label="Student name" name="name" value={form.name} onChange={updateForm} placeholder="Enter student name"/><Field label="Attendance (%)" name="attendance" value={form.attendance} onChange={updateForm} type="number" min="0" max="100"/><Field label="Study hours / day" name="study_hours" value={form.study_hours} onChange={updateForm} type="number" min="0" max="24" step="0.5"/><Field label="Previous marks (%)" name="previous_marks" value={form.previous_marks} onChange={updateForm} type="number" min="0" max="100"/><Field label="Assignment score (%)" name="assignment_score" value={form.assignment_score} onChange={updateForm} type="number" min="0" max="100"/><Field label="Sleep hours / day" name="sleep_hours" value={form.sleep_hours} onChange={updateForm} type="number" min="0" max="24" step="0.5"/></div>
                    {error && <div className="error-banner"><AlertCircle size={17}/><span>{error}</span></div>}
                    <div className="prediction-form-footer"><div className="privacy-note"><ShieldCheck size={16}/><span>Prediction is generated by your locally hosted model.</span></div><button type="submit" className="btn-primary predict-submit" disabled={loading}>{loading ? <><span className="button-spinner"/> Analyzing…</> : <><Sparkles size={17}/> Generate prediction</>}</button></div>
                  </form>
                </section>
                <section className="surface result-surface"><div className="surface-heading"><div><span className="mini-label">MODEL RESULT</span><h2>Prediction outcome</h2><p>Category and confidence returned by the API</p></div><div className="result-mark"><BrainCircuit size={21}/></div></div>
                  {prediction ? <PredictionResult prediction={prediction} form={form}/> : <div className="result-empty"><div className="result-empty-art"><BrainCircuit size={40}/><span><Sparkles size={17}/></span></div><h3>Ready when you are</h3><p>Fill in the student profile and run the model. Your result will appear here.</p><div className="result-checks"><span><Check size={14}/> Performance category</span><span><Check size={14}/> Model confidence</span><span><Check size={14}/> Focus recommendations</span></div></div>}
                </section>
              </div>
              <div className="info-note"><Zap size={17}/><p><b>How to read this result:</b> Confidence is the model's reported confidence for the predicted category. It is not a guarantee of a student's future academic outcome.</p></div>
            </>
          )}

          {page === "Analytics" && (
            <>
              <PageTitle eyebrow="PERFORMANCE INTELLIGENCE" title="Analytics & insights" subtitle="Attendance, study habits, academic marks and machine-learning performance categories." actions={<><button className="btn-secondary" onClick={exportHistory} disabled={!history.length}><Download size={16}/> Export history</button><button className="btn-primary" onClick={exportPerformanceReport} disabled={!students.length}><FileSpreadsheet size={16}/> Performance report</button><button className="btn-secondary" onClick={printPerformanceReport} disabled={!students.length}><FileSpreadsheet size={16}/> Print / Save PDF</button></>}/>
              <section className="kpi-grid analytics-kpis">
                <KpiCard icon={Activity} label="Total predictions" value={history.length} note="Saved model runs" tone="violet"/>
                <KpiCard icon={Award} label="Excellent results" value={excellent} note="Model classification" tone="mint"/>
                <KpiCard icon={TrendingUp} label="Average results" value={average} note="Model classification" tone="amber"/>
                <KpiCard icon={AlertCircle} label="At risk results" value={atRisk} note="Model classification" tone="rose"/>
              </section>
              <section className="analytics-focus-grid">
                <div className="surface chart-surface analytics-chart-card">
                  <div className="surface-heading"><div><span className="mini-label">MODEL OUTPUT</span><h2>Performance categories</h2><p>Excellent, Average and At Risk from saved model predictions</p></div></div>
                  {history.length ? <div className="analytics-donut"><ResponsiveContainer width="100%" height={260}><PieChart><Pie data={distribution} dataKey="value" nameKey="name" cx="50%" cy="46%" innerRadius={62} outerRadius={95} paddingAngle={4} label={({ name, percent }) => `${name} ${(percent * 100).toFixed(0)}%`}>{distribution.map(item => <Cell key={item.name} fill={item.color}/>)}</Pie><Tooltip/><Legend verticalAlign="bottom" height={30}/></PieChart></ResponsiveContainer></div> : <EmptyState icon={PieChartIcon} title="No prediction data" text="The category chart will appear after you run predictions." action={() => goTo("AI Prediction")} actionText="Run prediction"/>}
                </div>
                <div className="surface chart-surface analytics-chart-card">
                  <div className="surface-heading"><div><span className="mini-label">ATTENDANCE ANALYSIS</span><h2>Attendance distribution</h2><p>Registered students grouped by attendance percentage</p></div></div>
                  {students.length ? <ResponsiveContainer width="100%" height={260}><BarChart data={attendanceBands} margin={{ top: 12, right: 8, left: -18, bottom: 5 }}><CartesianGrid strokeDasharray="3 3" vertical={false} stroke="var(--chart-grid)"/><XAxis dataKey="band" axisLine={false} tickLine={false} tick={{ fill: "var(--muted)", fontSize: 11 }}/><YAxis allowDecimals={false} axisLine={false} tickLine={false} tick={{ fill: "var(--muted)", fontSize: 11 }}/><Tooltip/><Bar dataKey="count" name="Students" fill="#18a47a" radius={[7, 7, 0, 0]} barSize={30}/></BarChart></ResponsiveContainer> : <EmptyState icon={Users} title="No attendance records" text="Add students to analyse attendance." action={() => goTo("Students")} actionText="Add students"/>}
                </div>
                <div className="surface chart-surface analytics-chart-card">
                  <div className="surface-heading"><div><span className="mini-label">STUDY HOURS ANALYSIS</span><h2>Daily study pattern</h2><p>Students grouped by average study hours per day</p></div></div>
                  {students.length ? <ResponsiveContainer width="100%" height={260}><BarChart data={studyHourBands} margin={{ top: 12, right: 8, left: -18, bottom: 5 }}><CartesianGrid strokeDasharray="3 3" vertical={false} stroke="var(--chart-grid)"/><XAxis dataKey="band" axisLine={false} tickLine={false} tick={{ fill: "var(--muted)", fontSize: 11 }}/><YAxis allowDecimals={false} axisLine={false} tickLine={false} tick={{ fill: "var(--muted)", fontSize: 11 }}/><Tooltip/><Bar dataKey="count" name="Students" fill="#4c8dd8" radius={[7, 7, 0, 0]} barSize={30}/></BarChart></ResponsiveContainer> : <EmptyState icon={Clock3} title="No study-hour records" text="Add student records to analyse study habits." action={() => goTo("Students")} actionText="Add students"/>}
                </div>
                <div className="surface chart-surface analytics-chart-card">
                  <div className="surface-heading"><div><span className="mini-label">PREVIOUS MARKS ANALYSIS</span><h2>Academic marks bands</h2><p>Registered students grouped by previous marks percentage</p></div></div>
                  {students.length ? <ResponsiveContainer width="100%" height={260}><BarChart data={marksBands} margin={{ top: 12, right: 8, left: -18, bottom: 5 }}><CartesianGrid strokeDasharray="3 3" vertical={false} stroke="var(--chart-grid)"/><XAxis dataKey="band" axisLine={false} tickLine={false} tick={{ fill: "var(--muted)", fontSize: 11 }}/><YAxis allowDecimals={false} axisLine={false} tickLine={false} tick={{ fill: "var(--muted)", fontSize: 11 }}/><Tooltip/><Bar dataKey="count" name="Students" fill="#7775ef" radius={[7, 7, 0, 0]} barSize={34}/></BarChart></ResponsiveContainer> : <EmptyState icon={Award} title="No marks records" text="Add students to see academic marks distribution." action={() => goTo("Students")} actionText="Manage students"/>}
                </div>
              </section>
              <section className="surface analytics-summary-strip">
                <div><span>Average attendance</span><b>{students.length ? `${averageAttendance.toFixed(1)}%` : "—"}</b></div>
                <div><span>Average study time</span><b>{students.length ? `${averageStudyHours.toFixed(1)} hrs/day` : "—"}</b></div>
                <div><span>Average previous marks</span><b>{students.length ? `${averageMarks.toFixed(1)}%` : "—"}</b></div>
                <div><span>Students in directory</span><b>{students.length}</b></div>
              </section>
              <section className="surface insight-card"><div className="insight-card-icon"><Sparkles size={20}/></div><div><span className="mini-label">DATA SNAPSHOT</span><h2>What the current data shows</h2><p>{history.length ? `${history.length} prediction${history.length === 1 ? " has" : "s have"} been saved. ${excellent} are classified Excellent, ${average} Average and ${atRisk} At Risk by the model.` : "No AI predictions have been saved yet. Generate predictions to build your analytics."} {students.length ? `The directory contains ${students.length} student records. Average attendance is ${averageAttendance.toFixed(1)}%, average study time is ${averageStudyHours.toFixed(1)} hours per day and average previous marks are ${averageMarks.toFixed(1)}%.` : "Add students to build your academic profile."}</p></div><button className="btn-secondary" onClick={exportPerformanceReport} disabled={!students.length}>Download report <Download size={15}/></button></section>
            </>
          )}

          {page === "History" && (
            <>
              <PageTitle eyebrow="AUDIT TRAIL" title="Prediction history" subtitle="Review the results previously generated and saved by the model." actions={<button className="btn-secondary" onClick={exportHistory} disabled={!history.length}><Download size={16}/> Export CSV</button>}/>
              <div className="history-summary"><div><span className="history-summary-icon"><FileSpreadsheet size={18}/></span><div><b>{history.length}</b><span>Saved predictions</span></div></div><div><span className="history-summary-icon mint"><Check size={18}/></span><div><b>{excellent}</b><span>Excellent</span></div></div><div><span className="history-summary-icon amber"><Activity size={18}/></span><div><b>{average}</b><span>Average</span></div></div><div><span className="history-summary-icon rose"><AlertCircle size={18}/></span><div><b>{atRisk}</b><span>At Risk</span></div></div></div>
              <div className="surface directory-surface"><div className="directory-toolbar"><div><h2>All prediction records</h2><p>Every result returned from the prediction API</p></div><span className="records-count">{history.length} records</span></div>{history.length ? <div className="table-scroll"><table className="data-table"><thead><tr><th>STUDENT</th><th>PERFORMANCE</th><th>CONFIDENCE</th><th>DATE & TIME</th></tr></thead><tbody>{[...history].reverse().map((item, index) => <tr key={`${item.name}-${item.date}-${index}`}><td><div className="student-cell"><div className="student-avatar">{String(item.name || "S").charAt(0).toUpperCase()}</div><div><b>{item.name}</b><span>Prediction #{history.length - index}</span></div></div></td><td><PerformanceBadge value={item.student_performance}/></td><td><div className="confidence-cell"><div className="confidence-track"><span style={{ width: `${Math.max(0, Math.min(100, item.confidence))}%` }}/></div><b>{item.confidence.toFixed(1)}%</b></div></td><td className="date-cell">{formatDate(item.date)}</td></tr>)}</tbody></table></div> : <EmptyState icon={Clock3} title="No prediction history" text="Generated predictions will be saved here automatically." action={() => goTo("AI Prediction")} actionText="Run your first prediction"/>}</div>
            </>
          )}
        </div>
        <footer className="app-footer"><span>EduPredict <b>•</b> Student Performance Intelligence</span><span><span className="footer-status"/> Powered by your FastAPI + Random Forest backend</span></footer>
      </main>

      {toast && <div className="toast-message"><span><Check size={15}/></span>{toast}</div>}
    </div>
  );
}

function getInitials(name) {
  return String(name || "Admin").trim().split(/\s+/).slice(0, 2).map(part => part[0]?.toUpperCase() || "").join("") || "AD";
}

function AuthScreen({ mode, setMode, form, setForm, onSubmit, loading, error, showPassword, setShowPassword }) {
  const isRegister = mode === "register";
  const update = event => setForm(current => ({ ...current, [event.target.name]: event.target.value }));
  return (
    <div className="auth-page">
      <section className="auth-showcase">
        <div className="auth-brand"><div className="auth-brand-mark"><GraduationCap size={24}/></div><div><strong>EduPredict</strong><span>STUDENT INTELLIGENCE</span></div></div>
        <div className="auth-showcase-content">
          <span className="auth-eyebrow"><span/> ADMIN CONTROL CENTER</span>
          <h1>Turn academic data into <em>better decisions.</em></h1>
          <p>One secure workspace to manage student records, understand performance patterns and run AI-powered predictions.</p>
          <div className="auth-feature-list">
            <div><span><ShieldCheck size={17}/></span><div><b>Protected admin workspace</b><small>Sign in to access student and prediction data.</small></div></div>
            <div><span><BrainCircuit size={17}/></span><div><b>AI-powered performance insights</b><small>Use your trained Random Forest model.</small></div></div>
            <div><span><BarChart3 size={17}/></span><div><b>Clear academic analytics</b><small>Review records, trends and prediction history.</small></div></div>
          </div>
        </div>
        <div className="auth-showcase-visual" aria-hidden="true"><div className="auth-visual-grid"/><div className="auth-visual-orbit orbit-a"/><div className="auth-visual-orbit orbit-b"/><div className="auth-visual-core"><BrainCircuit size={42}/></div><div className="auth-float-card float-card-a"><span className="float-card-icon"><TrendingUp size={14}/></span><div><b>Academic insights</b><small>At a glance</small></div></div><div className="auth-float-card float-card-b"><span className="float-card-icon green"><ShieldCheck size={14}/></span><div><b>Admin workspace</b><small>Secure access</small></div></div></div>
        <div className="auth-showcase-footer"><span>EduPredict • Student Performance Prediction System</span><span>AI · Analytics · Student Success</span></div>
      </section>

      <main className="auth-main">
        <div className="auth-main-top"><span><span className="auth-status-dot"/> ADMIN PORTAL</span><span>Secure access</span></div>
        <section className="auth-card">
          <div className="auth-card-icon"><LockKeyhole size={21}/></div>
          <span className="auth-card-kicker">{isRegister ? "CREATE ADMIN ACCOUNT" : "WELCOME BACK"}</span>
          <h2>{isRegister ? "Create your workspace" : "Sign in to EduPredict"}</h2>
          <p className="auth-card-subtitle">{isRegister ? "Set up your administrator account to get started." : "Enter your administrator credentials to continue."}</p>

          <div className="auth-mode-switch">
            <button type="button" className={!isRegister ? "selected" : ""} onClick={() => setMode("login")}>Sign in</button>
            <button type="button" className={isRegister ? "selected" : ""} onClick={() => setMode("register")}>Register</button>
          </div>

          <form className="auth-form" onSubmit={onSubmit}>
            {isRegister && <label className="auth-field"><span>Full name</span><div className="auth-input-wrap"><UserRound size={17}/><input name="name" value={form.name} onChange={update} placeholder="Enter your full name" autoComplete="name" required minLength={2}/></div></label>}
            <label className="auth-field"><span>Email address</span><div className="auth-input-wrap"><Mail size={17}/><input name="email" type="email" value={form.email} onChange={update} placeholder="admin@college.edu" autoComplete="email" required/></div></label>
            <label className="auth-field"><span>Password</span><div className="auth-input-wrap"><LockKeyhole size={17}/><input name="password" type={showPassword ? "text" : "password"} value={form.password} onChange={update} placeholder={isRegister ? "At least 8 characters" : "Enter your password"} autoComplete={isRegister ? "new-password" : "current-password"} required minLength={isRegister ? 8 : 1}/><button type="button" className="password-toggle" onClick={() => setShowPassword(value => !value)} aria-label={showPassword ? "Hide password" : "Show password"}>{showPassword ? <EyeOff size={16}/> : <Eye size={16}/>}</button></div></label>
            {isRegister && <label className="auth-field"><span>Confirm password</span><div className="auth-input-wrap"><LockKeyhole size={17}/><input name="confirmPassword" type={showPassword ? "text" : "password"} value={form.confirmPassword} onChange={update} placeholder="Re-enter your password" autoComplete="new-password" required minLength={8}/></div></label>}
            {error && <div className="auth-error"><AlertCircle size={16}/><span>{error}</span></div>}
            <button className="auth-submit" type="submit" disabled={loading}>{loading ? <><span className="auth-button-spinner"/>{isRegister ? "Creating account…" : "Signing in…"}</> : <>{isRegister ? "Create administrator account" : "Sign in to dashboard"}<ArrowUpRight size={17}/></>}</button>
          </form>
          <div className="auth-security-note"><ShieldCheck size={15}/><span>Your account is authenticated by the EduPredict API. Passwords are stored as salted hashes.</span></div>
          <p className="auth-switch-copy">{isRegister ? "Already have an admin account?" : "New to EduPredict?"} <button type="button" onClick={() => setMode(isRegister ? "login" : "register")}>{isRegister ? "Sign in" : "Create an admin account"}</button></p>
        </section>
        <div className="auth-main-footer"><span>© {new Date().getFullYear()} EduPredict</span><span>Administrator access only</span></div>
      </main>
    </div>
  );
}

function KpiCard({ icon: Icon, label, value, note, tone }) {
  return <div className="kpi-card"><div className={`kpi-icon ${tone}`}><Icon size={19}/></div><div className="kpi-details"><span>{label}</span><strong>{value}</strong><small>{note}</small></div><div className="kpi-sparkline"><span/><span/><span/><span/><span/></div></div>;
}

function PageTitle({ eyebrow, title, subtitle, actions }) {
  return <div className="page-title-row"><div><span className="section-kicker">{eyebrow}</span><h1>{title}</h1><p>{subtitle}</p></div>{actions && <div className="page-title-actions">{actions}</div>}</div>;
}

function Field({ label, name, value, onChange, type = "text", required = true, ...props }) {
  return <label className="field"><span>{label}</span><input name={name} type={type} value={value ?? ""} onChange={onChange} required={required} {...props}/></label>;
}

function PerformanceBadge({ value }) {
  return <span className={`performance-badge ${performanceClass(value)}`}><i/>{value}</span>;
}

function performanceClass(value) {
  if (value === "Excellent") return "excellent";
  if (value === "Average") return "average";
  return "risk";
}

function StudentTable({ students, onEdit, onDelete, onAnalyze, onView }) {
  if (!students.length) return <EmptyState icon={Users} title="No students to show" text="Add a student or adjust your search and filters."/>;
  return <div className="table-scroll"><table className="data-table"><thead><tr><th>STUDENT</th><th>COURSE</th><th>ATTENDANCE</th><th>PREVIOUS MARKS</th><th>ACTIONS</th></tr></thead><tbody>{students.map(student => <tr key={student.id}><td><div className="student-cell"><div className="student-avatar">{String(student.name || "S").charAt(0).toUpperCase()}</div><div><b>{student.name}</b><span>{student.email || "No email provided"}</span></div></div></td><td><span className="course-pill">{student.course || "—"}</span></td><td><div className="metric-cell"><div className="metric-track"><span className={Number(student.attendance) < 65 ? "low" : ""} style={{ width: `${Math.max(0, Math.min(100, Number(student.attendance || 0)))}%` }}/></div><b>{Number(student.attendance || 0)}%</b></div></td><td><b className="marks-value">{Number(student.previous_marks || 0)}%</b></td><td><div className="table-actions"><button onClick={() => onView?.(student)} title="Open student dashboard"><Eye size={15}/></button><button onClick={() => onAnalyze(student)} title="Run prediction for this student"><BrainCircuit size={15}/></button><button onClick={() => onEdit(student)} title="Edit student"><Pencil size={15}/></button><button className="danger-action" onClick={() => onDelete(student.id)} title="Delete student"><Trash2 size={15}/></button></div></td></tr>)}</tbody></table></div>;
}

function PredictionResult({ prediction, form }) {
  const focus = [];
  if (Number(form.attendance) < 75) focus.push("Improve class attendance and consistency.");
  if (Number(form.previous_marks) < 60) focus.push("Review foundational topics and set weekly revision targets.");
  if (Number(form.assignment_score) < 65) focus.push("Submit assignments on time and practise applied questions.");
  if (Number(form.study_hours) < 3) focus.push("Build a steady daily study routine with focused sessions.");
  if (Number(form.sleep_hours) < 6) focus.push("Maintain a regular, healthy sleep schedule.");
  if (!focus.length) focus.push("Maintain the current study routine and review progress regularly.");
  return <div className="prediction-result"><div className={`prediction-result-icon ${performanceClass(prediction.student_performance)}`}><BrainCircuit size={30}/></div><span className="mini-label">PREDICTED PERFORMANCE</span><h3>{prediction.student_performance || "Result available"}</h3><p className="prediction-student-name">For {prediction.name}</p><div className="result-confidence"><div><span>Model confidence</span><b>{prediction.confidence.toFixed(1)}%</b></div><div className="result-confidence-track"><span style={{ width: `${Math.max(0, Math.min(100, prediction.confidence))}%` }}/></div></div><div className="recommendations"><b><Sparkles size={15}/> Suggested focus areas</b><ul>{focus.slice(0, 4).map(item => <li key={item}>{item}</li>)}</ul><small>These focus areas are simple rule-based suggestions from the entered values, not extra model outputs.</small></div></div>;
}

function EmptyState({ icon: Icon, title, text, action, actionText }) {
  return <div className="empty-state"><div className="empty-icon"><Icon size={22}/></div><b>{title}</b><p>{text}</p>{action && <button className="empty-action" onClick={action}>{actionText || "Continue"}<ArrowUpRight size={14}/></button>}</div>;
}

function PieChartIcon() {
  return <span className="pie-chart-symbol"><i/><i/><i/></span>;
}

export default App;
