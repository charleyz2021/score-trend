import { useEffect, useMemo, useRef, useState } from "react";
import ReactECharts from "echarts-for-react";
import type { ExamConfig, ExamId, MetricId } from "./importEngine";
import { importFiles, getEffectiveClass, toNumber } from "./importEngine";
import {
  AppShell,
  Container,
  Group,
  Stack,
  Card,
  Title,
  Text,
  Badge,
  Button,
  Divider,
  SimpleGrid,
  Select,
  type SelectProps,
  TextInput,
  Alert,
  Box,
  Paper,
} from "@mantine/core";

import { notifications } from "@mantine/notifications";
import {
  IconUpload,
  IconDownload,
  IconTrash,
  IconInfoCircle,
  IconAlertTriangle,
  IconChevronDown,
} from "@tabler/icons-react";

type RecordPoint = { examId: ExamId; examName: string; score: number | null };
type ExamStatus = "ok" | "warn" | "fail";

function statusText(s: ExamStatus) {
  if (s === "ok") return "正常";
  if (s === "warn") return "需检查";
  return "失败";
}
function statusColor(s: ExamStatus) {
  if (s === "ok") return "var(--mantine-color-green-7)";
  if (s === "warn") return "var(--mantine-color-yellow-7)";
  return "var(--mantine-color-red-7)";
}

const METRIC_LABEL: Record<MetricId, string> = {
  total: "总分",
  chinese: "语文",
  math: "数学",
  english: "英语",
  physics: "物理",
  chemistry: "化学",
  biology: "生物",
  politics: "政治",
  history: "历史",
  geography: "地理",
  classRank: "班级排名",
  schoolRank: "年级排名",
};

const METRIC_ORDER: MetricId[] = [
  "total",
  "chinese", "math", "english",
  "physics", "chemistry", "biology",
  "politics", "history", "geography",
  "classRank", "schoolRank",
];

// 规则：fatalErrors => fail；有 warnings / 总分未识别 / 同班重名 => warn；否则 ok
function getExamStatus(ex: ExamConfig, duplicateCountByExamId?: Map<ExamId, number>): ExamStatus {
  if (ex.fatalErrors?.length) return "fail";
  const dup = duplicateCountByExamId?.get(ex.id) ?? 0;
  if (dup > 0) return "warn";
  if (!ex.totalCol) return "warn";
  if (ex.warnings?.length) return "warn";
  return "ok";
}

const metricColOf = (ex: ExamConfig, metric: MetricId): string | undefined => {
  if (metric === "total") return ex.metricCols?.total || ex.totalCol || undefined;
  return ex.metricCols?.[metric] || undefined;
};

const hasUsableClassCol = (ex: ExamConfig) => {
  if (!ex.classCol) return false;
  const N = 300;
  for (const r of ex.rows.slice(0, N)) {
    const v = String(r[ex.classCol] ?? "").trim();
    if (v) return true; // 只要出现过一个非空班级，就算“可用”
  }
  return false; // 表头有，但全空
};

export default function App() {
  const [exams, setExams] = useState<ExamConfig[]>([]);
  const [activeExamId, setActiveExamId] = useState<ExamId>("");
  const [student, setStudent] = useState<string>("");

  // 班级筛选
  const [classFilter, setClassFilter] = useState<string>("全校");
  const [includeUnknown, setIncludeUnknown] = useState<boolean>(false);
  const [fileLabel, setFileLabel] = useState<string>("");
  const [editOverride, setEditOverride] = useState(false);
  const [overrideDraft, setOverrideDraft] = useState("");
  const [pendingClassFilter, setPendingClassFilter] = useState<string>("");
  const [zoomRange, setZoomRange] = useState<{ start: number; end: number }>({ start: 0, end: 100 });

  const fileRef = useRef<HTMLInputElement | null>(null);

  // ✅ 默认指标
  const [metric, setMetric] = useState<MetricId>("total");

  async function onFilesChange(e: React.ChangeEvent<HTMLInputElement>) {
    const files = Array.from(e.target.files ?? []);
    if (!files.length) return;

    try {
      setFileLabel(files.map(f => f.name).join("、"));

      const MAX_MB = 10;
      for (const f of files) {
        if (f.size > MAX_MB * 1024 * 1024) {
          alert(`文件过大（>${MAX_MB}MB）：${f.name}`);
          return;
        }
      }

      const all = await importFiles(files);

      setExams(all);
      setActiveExamId(all[0]?.id ?? "");
      setStudent("");
      setMetric("total");

      // 默认班级
      const count = new Map<string, number>();
      for (const ex of all) {
        const c = getEffectiveClass(ex);
        if (!c || c === "全校" || c === "未知班级") continue;
        count.set(c, (count.get(c) ?? 0) + 1);
      }
      let best = "全校";
      let bestV = 0;
      for (const [c, v] of count.entries()) {
        if (v > bestV) { bestV = v; best = c; }
      }
      setClassFilter("全校");
      setPendingClassFilter(best);
      setIncludeUnknown(false);
    } catch (err: any) {
      // eslint-disable-next-line no-console
      console.error("[importFiles failed]", err);
      alert(`导入失败：${err?.message ?? err}`);
    } finally {
      if (fileRef.current) fileRef.current.value = "";
    }
  }

  const onClearAll = () => {
    setExams([]);
    setActiveExamId("");
    setStudent("");
    setClassFilter("全校");
    setIncludeUnknown(false);
    setFileLabel("");
    setMetric("total");
    if (fileRef.current) fileRef.current.value = "";
  };

  function updateExam(examId: ExamId, patch: Partial<ExamConfig>) {
    setExams(prev => prev.map(x => (x.id === examId ? { ...x, ...patch } : x)));
  }

  const activeExam = useMemo(
    () => exams.find(e => e.id === activeExamId) ?? null,
    [exams, activeExamId]
  );

  const classOptions = useMemo(() => {
    const set = new Set<string>();
    let hasUnknown = false;

    for (const ex of exams ?? []) {
      if (ex.classCol && hasUsableClassCol(ex)) {
        if (!ex.nameCol) continue;
        for (const r of ex.rows) {
          const cls = String(r[ex.classCol] ?? "").trim();
          if (cls) set.add(cls);
          else hasUnknown = true;
        }
        continue;
      }

      const c = String(getEffectiveClass(ex) ?? "").trim();
      if (!c || c === "全校" || c === "未知班级") hasUnknown = true;
      else set.add(c);
    }

    const arr = Array.from(set);
    arr.sort((a, b) => {
      const na = Number(a), nb = Number(b);
      const aNum = Number.isFinite(na), bNum = Number.isFinite(nb);
      if (aNum && bNum) return na - nb;
      if (aNum) return -1;
      if (bNum) return 1;
      return a.localeCompare(b, "zh-Hans-CN");
    });

    const out = ["全校", ...arr];
    if (hasUnknown) out.push("未知班级");
    return out;
  }, [exams]);

  // ✅ 每个考试：name -> 当前 metric 数值
  const examScoreMaps = useMemo(() => {
    const maps = new Map<ExamId, Map<string, number | null>>();

    for (const ex of exams) {
      const m = new Map<string, number | null>();

      if (ex.fatalErrors?.length) { maps.set(ex.id, m); continue; }
      if (!ex.nameCol) { maps.set(ex.id, m); continue; }

      for (const r of ex.rows) {
        const name = String(r[ex.nameCol] ?? "").trim();
        if (!name) continue;

        // 班级过滤
        if (classFilter !== "全校") {
          if (ex.classCol && hasUsableClassCol(ex)) {
            const rowCls = String(r[ex.classCol!] ?? "").trim();
            if (rowCls !== classFilter) continue;
          } else {
            const eff = (ex.overrideClass?.trim() || ex.inferredClass || "未知班级");
            if (eff !== classFilter) continue;
          }
        }

        const colKey = metricColOf(ex, metric);
        const sc = colKey ? toNumber(r[colKey]) : null;

        if (sc === null) {
          if (!m.has(name)) m.set(name, null);
        } else {
          m.set(name, sc);
        }
      }

      maps.set(ex.id, m);
    }

    return maps;
  }, [exams, classFilter, metric]);

  const filteredExams = useMemo(() => {
    if (!exams?.length) return [];
    if (classFilter === "全校") return exams;

    return exams.filter(ex => {
      if (hasUsableClassCol(ex)) return true;
      const cls = getEffectiveClass(ex);
      if (cls === classFilter) return true;
      if (includeUnknown && cls === "未知班级") return true;
      return false;
    });
  }, [exams, classFilter, includeUnknown]);

  const studentOptions = useMemo(() => {
    const s = new Set<string>();

    for (const ex of filteredExams) {
      if (!ex.nameCol) continue;

      if (ex.classCol && hasUsableClassCol(ex)) {
        for (const r of ex.rows) {
          const name = String(r[ex.nameCol] ?? "").trim();
          if (!name) continue;

          if (classFilter === "全校") { s.add(name); continue; }

          const cls = String(r[ex.classCol] ?? "").trim() || "未知班级";
          if (cls === classFilter) s.add(name);
        }
        continue;
      }

      const sheetClassRaw = String(getEffectiveClass(ex) ?? "").trim();
      const sheetClass = sheetClassRaw && sheetClassRaw !== "全校" ? sheetClassRaw : "未知班级";
      if (classFilter !== "全校" && sheetClass !== classFilter) continue;

      for (const r of ex.rows) {
        const name = String(r[ex.nameCol] ?? "").trim();
        if (name) s.add(name);
      }
    }

    return Array.from(s).sort((a, b) => a.localeCompare(b, "zh-Hans-CN"));
  }, [filteredExams, classFilter]);

  // ✅ 当前学生在“所有考试中”是否存在某 metric 的任何数值（用于置灰）
  const studentHasMetricData = useMemo(() => {
    const has = new Map<MetricId, boolean>();

    const allMetrics: MetricId[] = [
      "total",
      "chinese", "math", "english",
      "physics", "chemistry", "biology",
      "history", "geography", "politics",
      "classRank", "schoolRank",
    ];

    for (const mid of allMetrics) {
      if (!student) { has.set(mid, true); continue; } // 未选学生：不置灰
      let ok = false;
      for (const ex of exams) {
        if (ex.fatalErrors?.length) continue;
        if (!ex.nameCol) continue;

        const colKey = metricColOf(ex, mid);
        if (!colKey) continue;

        // 找到该学生的行
        // （为了省事：扫行，量不大）
        for (const r of ex.rows) {
          const name = String(r[ex.nameCol] ?? "").trim();
          if (!name || name !== student) continue;
          const v = toNumber(r[colKey]);
          if (v !== null) { ok = true; break; }
        }
        if (ok) break;
      }
      has.set(mid, ok);
    }

    return has;
  }, [student, exams]);

  const metricOptions = useMemo(() => {
    if (!exams.length) return [{ value: "total", label: "总分" }];

    // 1) 先统计：全局有哪些 metric 列存在（任意一场出现就算存在）
    const exists = new Set<MetricId>();
    for (const ex of exams) {
      const m = ex.metricCols || {};
      for (const k of METRIC_ORDER) {
        if (k === "total") {
          if (m.total || ex.totalCol) exists.add("total");
        } else {
          if (m[k]) exists.add(k);
        }
      }
    }

    // 2) 再统计：对“当前学生 + 当前班级筛选”，每个 metric 是否真的有数值
    const hasValue = new Set<MetricId>();
    if (student) {
      for (const ex of exams) {
        if (ex.fatalErrors?.length) continue;
        if (!ex.nameCol) continue;

        // 找到该学生行（同名多条时，只要有一条能转成数字就算有）
        for (const r of ex.rows) {
          const name = String(r[ex.nameCol] ?? "").trim();
          if (name !== student) continue;

          // 班级筛选（与你 examScoreMaps 同逻辑）
          if (classFilter !== "全校") {
            if (ex.classCol && hasUsableClassCol(ex)) {
              const rowCls = String(r[ex.classCol] ?? "").trim() || "未知班级";
              if (rowCls !== classFilter) continue;
            } else {
              const eff = (ex.overrideClass?.trim() || ex.inferredClass || "未知班级");
              if (eff !== classFilter) continue;
            }
          }

          for (const k of METRIC_ORDER) {
            const colKey = (k === "total")
              ? (ex.metricCols?.total || ex.totalCol)
              : ex.metricCols?.[k];

            if (!colKey) continue;
            const n = toNumber(r[colKey]);
            if (typeof n === "number") hasValue.add(k);
          }
        }
      }
    }

    // 3) 生成 Select data：存在但该学生没数据 -> 灰 + 不可选
    return METRIC_ORDER
      .filter(k => exists.has(k)) // 全局不存在的就不显示
      .map(k => {
        const noData = student ? !hasValue.has(k) : false;
        return {
          value: k,
          label: noData ? `${METRIC_LABEL[k]}（无数据）` : METRIC_LABEL[k],
          disabled: noData,
        };
      });
  }, [exams, student, classFilter]);

  // ✅ 如果当前 metric 变成“被置灰/不存在”，自动回退总分
  useEffect(() => {
    const opt = metricOptions.find(o => o.value === metric);
    if (!opt || (opt as any).disabled) setMetric("total");
  }, [metricOptions, metric]);

  // 学生时间序列
  const studentSeries = useMemo(() => {
    if (!student) return [];

    const out: RecordPoint[] = [];
    for (const ex of exams) {
      if (ex.fatalErrors.length) continue;
      const m = examScoreMaps.get(ex.id);
      const sc = m?.get(student) ?? null;
      out.push({ examId: ex.id, examName: ex.examName, score: sc });
    }
    return out;
  }, [student, exams, examScoreMaps]);

  const chartOption = useMemo(() => {
    if (!student) return null;

    const x = studentSeries.map((p) => p.examName);
    const y = studentSeries.map((p) => p.score);

    const hasAny = y.some((v) => typeof v === "number");
    if (!hasAny) return null;

    const start = zoomRange.start ?? 0;
    const end = zoomRange.end ?? 100;

    const span = Math.max(1, end - start);
    const axisFontSize = Math.max(10, Math.min(18, Math.round(10 + (100 - span) * 0.08)));

    const rotate = 35;
    const isRank = metric === "classRank" || metric === "schoolRank";

    return {
      title: { text: `${student} · ${METRIC_LABEL[metric]}波动` },
      tooltip: { trigger: "axis" },
      grid: { left: 50, right: 20, top: 60, bottom: 150 },

      xAxis: {
        type: "category",
        data: x,
        axisLabel: {
          rotate,
          interval: 0,
          fontSize: axisFontSize,
          margin: 16,
          align: "right",
          hideOverlap: false,
        },
      },

      yAxis: { type: "value", inverse: isRank },

      dataZoom: [
        { type: "inside", xAxisIndex: 0, start, end },
        {
          type: "slider",
          xAxisIndex: 0,
          start,
          end,
          bottom: 42,
          height: 16,
          showDetail: false,
          labelFormatter: () => "",
        },
      ],

      series: [
        {
          name: METRIC_LABEL[metric],
          type: "line",
          data: y,
          smooth: true,
          connectNulls: false,
          showSymbol: true,
          symbolSize: 6,
        },
      ],
    };
  }, [student, studentSeries, zoomRange, metric]);

  // 同班重名提示
  const { duplicateNameWarnings, duplicateCountByExamId } = useMemo(() => {
    const warnings: string[] = [];
    const m = new Map<ExamId, number>();

    for (const ex of exams) {
      if (ex.fatalErrors?.length) continue;
      if (!ex.nameCol) continue;
      if (ex.idCol) continue;
      const sheetClass = getEffectiveClass(ex);
      if (sheetClass === "全校") continue;

      const seen = new Map<string, number>();
      for (const r of ex.rows) {
        const name = String(r[ex.nameCol] ?? "").trim();
        if (!name) continue;
        seen.set(name, (seen.get(name) ?? 0) + 1);
      }
      const dups = Array.from(seen.entries()).filter(([, c]) => c >= 2).map(([n]) => n);
      if (dups.length) {
        m.set(ex.id, dups.length);
        warnings.push(
          `考试「${ex.examName}」检测到班级「${sheetClass}」内重名：${dups.slice(0, 6).join("、")}${dups.length > 6 ? "…" : ""}。建议：在表中补充学生ID/学号或在姓名中添加区分标识（如“张三-1/张三-2”）。`
        );

      }
    }

    return { duplicateNameWarnings: warnings, duplicateCountByExamId: m };
  }, [exams]);

  const hasUnknownStudents = useMemo(() => {
    for (const ex of exams ?? []) {
      if (!ex.classCol) continue;
      if (!ex.nameCol) continue;
      for (const r of ex.rows) {
        const cls = String(r[ex.classCol] ?? "").trim();
        const name = String(r[ex.nameCol] ?? "").trim();
        if (name && !cls) return true;
      }
    }
    return (exams ?? []).some(ex => getEffectiveClass(ex) === "未知班级");
  }, [exams]);

  const examSelectData = useMemo(() => {
    return exams.map((ex) => ({
      value: ex.id,
      label: ex.examName,
      status: getExamStatus(ex, duplicateCountByExamId),
    }));
  }, [exams, duplicateCountByExamId]);

  const globalStatus: ExamStatus = useMemo(() => {
    if (!exams.length) return "ok";
    let hasWarn = false;
    for (const ex of exams) {
      const s = getExamStatus(ex, duplicateCountByExamId);
      if (s === "fail") return "fail";
      if (s === "warn") hasWarn = true;
    }
    return hasWarn ? "warn" : "ok";
  }, [exams, duplicateCountByExamId]);

  const globalStatusUI = useMemo(() => {
    if (globalStatus === "fail") return { text: "失败", color: "red" as const };
    if (globalStatus === "warn") return { text: "需检查", color: "yellow" as const };
    return { text: "正常", color: "green" as const };
  }, [globalStatus]);

  const formatErrors = useMemo(() => {
    const out: string[] = [];
    for (const ex of exams) {
      for (const e of ex.fatalErrors ?? []) out.push(`考试「${ex.examName}」：${e}`);
    }
    return out;
  }, [exams]);

  const classSelectData = useMemo(
    () => (classOptions ?? [])
      .filter((c) => c !== undefined && c !== null && String(c).trim() !== "")
      .map((c) => ({ value: String(c), label: String(c) })),
    [classOptions]
  );

  const activeExamStatus: ExamStatus = useMemo(() => {
    if (!activeExam) return "ok";
    return getExamStatus(activeExam, duplicateCountByExamId);
  }, [activeExam, duplicateCountByExamId]);

  // activeExamId 必须在 exams 里
  useEffect(() => {
    if (exams.length) {
      if (!exams.some((ex) => ex.id === activeExamId)) setActiveExamId(exams[0].id);
    } else {
      if (activeExamId) setActiveExamId("");
    }
  }, [exams, activeExamId]);

  // classFilter 必须在 classOptions 里
  useEffect(() => {
    if (!classOptions.length) {
      if (classFilter !== "全校") setClassFilter("全校");
      return;
    }
    if (!classOptions.includes(classFilter)) setClassFilter("全校");
  }, [classOptions, classFilter]);

  // 导入后恢复 best
  useEffect(() => {
    if (!pendingClassFilter) return;
    if (classOptions.includes(pendingClassFilter)) {
      setClassFilter(pendingClassFilter);
      setPendingClassFilter("");
    }
  }, [pendingClassFilter, classOptions]);

  // 包含未知班级开关
  useEffect(() => {
    if (!hasUnknownStudents && includeUnknown) setIncludeUnknown(false);
  }, [hasUnknownStudents, includeUnknown]);

  // 切换时重置覆盖输入框
  useEffect(() => {
    if (!activeExam) return;
    setEditOverride(false);
    setOverrideDraft(activeExam.overrideClass || "");
  }, [activeExamId, activeExam]);

  useEffect(() => {
    if (student && !studentOptions.includes(student)) setStudent("");
  }, [student, studentOptions]);

  const renderExamOption: SelectProps["renderOption"] = ({ option }) => {
    const s = (option as any).status as ExamStatus;
    return (
      <Group justify="space-between" wrap="nowrap" w="100%">
        <Text size="sm" style={{ flex: 1, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
          {option.label}
        </Text>
        <Text size="sm" fw={600} style={{ color: statusColor(s), whiteSpace: "nowrap" }}>
          {statusText(s)}
        </Text>
      </Group>
    );
  };

  const onDownloadTemplate = () => {
    const a = document.createElement("a");
    a.href = "/template.xlsx";
    a.download = "成绩模板.xlsx";
    document.body.appendChild(a);
    a.click();
    a.remove();
  };

  const saveOverrideClass = () => {
    if (!activeExam) return;
    const v = overrideDraft.trim();

    if (v && !/^\d{1,3}$/.test(v) && v !== "未知班级") {
      notifications.show({
        color: "red",
        title: "输入不合法",
        message: "班级建议填写数字班号（如 3），或留空使用自动推断。",
      });
      return;
    }

    updateExam(activeExam.id, { overrideClass: v });
    setOverrideDraft(v);
    setEditOverride(false);
  };

  const clearOverrideClass = () => {
    if (!activeExam) return;
    updateExam(activeExam.id, { overrideClass: "" });
    setOverrideDraft("");
    setEditOverride(false);
  };

  const studentHasSelectedMetricData = useMemo(() => {
    if (!student) return true;
    return studentHasMetricData.get(metric) ?? true;
  }, [student, metric, studentHasMetricData]);

  return (
    <AppShell
      header={{ height: 64 }}
      padding="md"
      styles={{ main: { background: "transparent" } }}
    >
      <AppShell.Header style={{ background: "white", borderBottom: "1px solid var(--mantine-color-gray-2)" }}>
        <Container size={1120} h="100%">
          <Group justify="space-between" h="100%">
            <Stack gap={2}>
              <Title fw={600} order={4} m={0}>成绩可视化助手</Title>
              <Text size="xs" c="dimmed">多文件/多Sheet 导入 → 班级筛选 → 学生分数趋势</Text>
            </Stack>

            <Group gap="sm">
              <Badge variant="light" size="lg">
                已导入考试：{exams.length}
              </Badge>

              {exams.length ? (
                <Badge variant="light" size="lg" color={globalStatusUI.color}>
                  {globalStatusUI.text}
                </Badge>
              ) : null}
            </Group>
          </Group>
        </Container>
      </AppShell.Header>

      <AppShell.Main className="skyMain" style={{ display: "flex", alignItems: "flex-start" }}>
        <Container size={1120} style={{ width: "100%" }}>
          <Box
            style={{
              minHeight: "calc(100vh - 64px - 32px)",
              display: "flex",
              alignItems: "flex-start",
              justifyContent: "center",
              paddingTop: 28,
            }}
          >
            <Box style={{ width: "100%" }}>
              {/* 顶部工具条 */}
              <Card withBorder radius="lg" p="md" mb="md" shadow="sm">
                <Group justify="space-between" wrap="wrap">
                  <Group gap="sm" wrap="wrap">
                    <input
                      ref={fileRef}
                      type="file"
                      accept=".xls,.xlsx,.csv"
                      multiple
                      onChange={onFilesChange}
                      style={{ display: "none" }}
                    />

                    <Button variant="light" leftSection={<IconDownload size={18} />} onClick={onDownloadTemplate}>
                      下载模板
                    </Button>

                    <Button leftSection={<IconUpload size={18} />} onClick={() => fileRef.current?.click()}>
                      选择文件
                    </Button>

                    <Badge variant="outline" color={fileLabel ? "blue" : "gray"}>
                      {fileLabel ? `已选择：${fileLabel}` : "未选择文件"}
                    </Badge>
                  </Group>

                  <Button
                    color="red"
                    variant="light"
                    leftSection={<IconTrash size={18} />}
                    onClick={onClearAll}
                    disabled={!exams.length}
                  >
                    清空
                  </Button>
                </Group>
              </Card>

              {/* 全局错误/警告 */}
              {formatErrors.length ? (
                <Alert
                  mb="md"
                  color="red"
                  variant="light"
                  icon={<IconAlertTriangle size={18} />}
                  title={`发现 ${formatErrors.length} 条格式错误（对应考试不可用），请按照模板修改`}
                >
                  <Stack gap={4}>
                    {formatErrors.slice(0, 6).map((t, i) => (
                      <Text key={i} size="sm">• {t}</Text>
                    ))}
                    {formatErrors.length > 6 ? <Text size="sm">• …</Text> : null}
                  </Stack>
                </Alert>
              ) : null}

              {duplicateNameWarnings.length ? (
                <Alert
                  color="yellow"
                  variant="light"
                  title="发现同班重名（需检查）"
                  icon={<IconAlertTriangle size={18} />}
                  mb="md"
                >
                  {duplicateNameWarnings.slice(0, 4).map((w, i) => (
                    <Text key={i} size="sm">• {w}</Text>
                  ))}
                  {duplicateNameWarnings.length > 4 ? <Text size="sm">• …</Text> : null}
                </Alert>
              ) : null}

              {exams.length ? (
                <SimpleGrid cols={{ base: 1, md: 2 }} spacing="md">
                  {/* 左：考试配置 */}
                  <Card withBorder radius="lg" p="md" shadow="sm">
                    <Group justify="space-between" mb="xs">
                      <Text fw={700}>考试信息</Text>
                      <Text size="xs" c="dimmed">可选，不影响直接看图</Text>
                    </Group>
                    <Divider mb="md" />

                    <Stack gap="md">
                      <Select
                        label="考试"
                        value={activeExamId}
                        onChange={(v) => setActiveExamId(v || "")}
                        data={examSelectData as any}
                        renderOption={renderExamOption}
                        searchable
                        nothingFoundMessage="未找到"
                        rightSection={
                          <Group gap={6} wrap="nowrap">
                            <Text size="sm" fw={600} style={{ color: statusColor(activeExamStatus) }}>
                              {statusText(activeExamStatus)}
                            </Text>
                            <IconChevronDown size={16} />
                          </Group>
                        }
                        rightSectionWidth={72}
                      />

                      {activeExam?.fatalErrors?.length ? (
                        <Alert
                          color="red"
                          variant="light"
                          icon={<IconAlertTriangle size={18} />}
                          title="该考试导入失败"
                          mt="sm"
                        >
                          <Stack gap={4}>
                            {activeExam.fatalErrors.slice(0, 4).map((x, i) => (
                              <Text key={i} size="sm">• {x}</Text>
                            ))}
                            {activeExam.fatalErrors.length > 4 ? <Text size="sm">• …</Text> : null}
                            <Text size="xs" c="dimmed" mt={6}>
                              建议：按模板检查表头（姓名列每sheet只能有一列、总分列需为数值等）。
                            </Text>
                          </Stack>
                        </Alert>
                      ) : null}

                      {activeExam?.warnings?.length ? (
                        <Alert
                          color="yellow"
                          variant="light"
                          icon={<IconInfoCircle size={18} />}
                          title="提示"
                        >
                          <Stack gap={4}>
                            {activeExam.warnings.map((x, i) => (
                              <Text key={i} size="sm">• {x}</Text>
                            ))}
                          </Stack>
                        </Alert>
                      ) : null}

                      {activeExam ? (
                        <>
                          <TextInput
                            label="考试名（x轴显示，可修改）"
                            value={activeExam.examName}
                            onChange={(e) => updateExam(activeExam.id, { examName: e.currentTarget.value })}
                          />

                          <Group gap="xs">
                            <Badge color={activeExam?.nameCol ? "green" : "red"} variant="light">
                              姓名列：{activeExam?.nameCol ? "已识别" : "缺失"}
                            </Badge>
                            <Badge color={activeExam?.totalCol ? "green" : "yellow"} variant="light">
                              总分列：{activeExam?.totalCol ? "已识别" : "未识别"}
                            </Badge>
                          </Group>

                          {activeExam && !hasUsableClassCol(activeExam) ? (
                            <Paper withBorder radius="md" p="md">
                              <Group justify="space-between" align="center" mb={6}>
                                <Text fw={700} size="sm">该表班级归属</Text>
                                <Badge variant="light" color={activeExam.overrideClass ? "violet" : "blue"}>
                                  {activeExam.overrideClass
                                    ? `已覆盖：${activeExam.overrideClass}`
                                    : `自动推断：${getEffectiveClass(activeExam)} `}
                                </Badge>
                              </Group>

                              <Text size="xs" c="dimmed" mb="sm">
                                仅当该 sheet 没有“班级”列且推断不准时，才需要覆盖。正确就不用改。
                              </Text>

                              {!editOverride ? (
                                <Group justify="space-between" align="center">
                                  <Text size="sm">
                                    当前用于筛选： <b>{activeExam.overrideClass || getEffectiveClass(activeExam)}</b>
                                  </Text>

                                  <Group gap="xs">
                                    <Button variant="light" onClick={() => setEditOverride(true)}>修改</Button>
                                    {activeExam.overrideClass ? (
                                      <Button color="red" variant="light" onClick={clearOverrideClass}>
                                        清除覆盖
                                      </Button>
                                    ) : null}
                                  </Group>
                                </Group>
                              ) : (
                                <Stack gap="xs">
                                  <TextInput
                                    value={overrideDraft}
                                    onChange={(e) => setOverrideDraft(e.currentTarget.value)}
                                    placeholder="输入班级号（如 3）"
                                    autoFocus
                                  />
                                  <Group justify="flex-end" gap="xs">
                                    <Button
                                      variant="light"
                                      onClick={() => {
                                        setEditOverride(false);
                                        setOverrideDraft(activeExam.overrideClass || "");
                                      }}
                                    >
                                      取消
                                    </Button>
                                    {activeExam.overrideClass ? (
                                      <Button color="red" variant="light" onClick={clearOverrideClass}>
                                        清除覆盖
                                      </Button>
                                    ) : null}
                                    <Button onClick={saveOverrideClass}>确认</Button>
                                  </Group>
                                </Stack>
                              )}
                            </Paper>
                          ) : null}
                        </>
                      ) : null}
                    </Stack>
                  </Card>

                  {/* 右：筛选 + 图表 */}
                  <Card withBorder radius="lg" p="md" shadow="sm" miw={0}>
                    <Group justify="space-between" mb="xs">
                      <Text fw={700}>筛选与图表</Text>
                    </Group>
                    <Divider mb="md" />

                    <SimpleGrid cols={{ base: 1, sm: 3 }} spacing="md">
                      <Select
                        label="班级筛选"
                        value={classFilter || null}
                        onChange={(v) => setClassFilter(v || "全校")}
                        data={classSelectData}
                      />

                      <Select
                        label="选择学生"
                        value={student || null}
                        onChange={(v) => setStudent(v || "")}
                        data={studentOptions.map((s) => ({ value: s, label: s }))}
                        searchable
                        clearable
                        nothingFoundMessage="未找到"
                        placeholder="请选择"
                      />

                      <Select
                        label="波动指标"
                        value={metric}
                        onChange={(v) => setMetric((v as MetricId) || "total")}
                        data={metricOptions as any}
                        nothingFoundMessage="该数据未识别到可用指标"
                      />
                    </SimpleGrid>

                    <Box mt="md">
                      {student ? (
                        studentHasSelectedMetricData ? (
                          chartOption ? (
                            <>
                              <Text size="xs" c="dimmed" mb={6}>
                                下方缩放条：拖动两端可放大/缩小时间范围；按住中间可整体平移
                              </Text>
                              <ReactECharts
                                option={chartOption as any}
                                style={{ height: 520 }}
                                onEvents={{
                                  datazoom: (params: any) => {
                                    const b = params?.batch?.[0];
                                    if (b && typeof b.start === "number" && typeof b.end === "number") {
                                      setZoomRange({ start: b.start, end: b.end });
                                    }
                                  },
                                }}
                              />
                            </>
                          ) : (
                            <Alert variant="light" color="blue" icon={<IconInfoCircle size={18} />}>
                              当前指标在该学生身上没有可绘制的数值点（可能全部为空或未导出）。
                            </Alert>
                          )
                        ) : (
                          <Alert variant="light" color="yellow" icon={<IconInfoCircle size={18} />}>
                            该学生在「{METRIC_LABEL[metric]}」没有可用数据（所以该项已置灰不可选）。
                          </Alert>
                        )
                      ) : (
                        <Alert variant="light" color="blue" icon={<IconInfoCircle size={18} />}>
                          选择学生后显示趋势。
                        </Alert>
                      )}
                    </Box>
                  </Card>
                </SimpleGrid>
              ) : (
                <Card withBorder radius="lg" p="xl">
                  <Title order={5} mb={6}>开始使用</Title>
                  <Text c="dimmed" size="sm">
                    点击“下载模板”查看成绩单格式说明，点击“选择文件”导入 Excel/CSV。
                  </Text>
                </Card>
              )}
            </Box>
          </Box>
        </Container>
      </AppShell.Main>
    </AppShell>
  );
}
