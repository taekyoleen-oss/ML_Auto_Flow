import PptxGenJS from "pptxgenjs";
import { CanvasModule, Connection, ModuleStatus, ModuleType } from "../types";

// ── 색상/테마 상수 ──────────────────────────────────────────────
const COLORS = {
  primary: "1E40AF",    // 파란색
  accent: "0891B2",     // 청록색
  success: "16A34A",    // 초록색
  warning: "D97706",    // 주황색
  error: "DC2626",      // 빨간색
  white: "FFFFFF",
  darkText: "1E293B",
  mutedText: "64748B",
  lightBg: "F8FAFC",
  border: "E2E8F0",
};

// ── 모듈 타입 한국어 이름 ──────────────────────────────────────
const MODULE_TYPE_LABELS: Partial<Record<ModuleType, string>> = {
  [ModuleType.LoadData]: "데이터 로드",
  [ModuleType.Statistics]: "통계 분석",
  [ModuleType.SelectData]: "데이터 선택",
  [ModuleType.DataFiltering]: "데이터 필터링",
  [ModuleType.ColumnPlot]: "컬럼 시각화",
  [ModuleType.OutlierDetector]: "이상값 탐지",
  [ModuleType.HypothesisTesting]: "가설 검정",
  [ModuleType.NormalityChecker]: "정규성 검정",
  [ModuleType.Correlation]: "상관관계 분석",
  [ModuleType.HandleMissingValues]: "결측치 처리",
  [ModuleType.TransformData]: "데이터 변환",
  [ModuleType.EncodeCategorical]: "범주형 인코딩",
  [ModuleType.ScalingTransform]: "스케일링 변환",
  [ModuleType.SplitData]: "데이터 분할",
  [ModuleType.Join]: "데이터 조인",
  [ModuleType.Concat]: "데이터 연결",
  [ModuleType.LinearRegression]: "선형 회귀",
  [ModuleType.LogisticRegression]: "로지스틱 회귀",
  [ModuleType.DecisionTree]: "의사결정나무",
  [ModuleType.RandomForest]: "랜덤 포레스트",
  [ModuleType.NeuralNetwork]: "신경망",
  [ModuleType.SVM]: "SVM",
  [ModuleType.LDA]: "LDA",
  [ModuleType.NaiveBayes]: "나이브 베이즈",
  [ModuleType.KNN]: "K-최근접 이웃",
  [ModuleType.TrainModel]: "모델 학습",
  [ModuleType.ScoreModel]: "모델 스코어링",
  [ModuleType.EvaluateModel]: "모델 평가",
  [ModuleType.KMeans]: "K-Means 클러스터링",
  [ModuleType.PCA]: "주성분 분석 (PCA)",
  [ModuleType.TrainClusteringModel]: "클러스터링 모델 학습",
  [ModuleType.ClusteringData]: "클러스터링 데이터",
  [ModuleType.OLSModel]: "OLS 회귀 모델",
  [ModuleType.LogisticModel]: "로지스틱 모델",
  [ModuleType.PoissonModel]: "포아송 모델",
  [ModuleType.EvaluateStat]: "통계 모델 평가",
  [ModuleType.VIFChecker]: "VIF 검정",
  [ModuleType.DiversionChecker]: "분산 분석",
};

function getModuleLabel(type: ModuleType): string {
  return MODULE_TYPE_LABELS[type] ?? type;
}

function getStatusLabel(status: ModuleStatus): string {
  switch (status) {
    case ModuleStatus.Success: return "성공";
    case ModuleStatus.Error:   return "오류";
    case ModuleStatus.Running: return "실행 중";
    default:                   return "대기";
  }
}

function getStatusColor(status: ModuleStatus): string {
  switch (status) {
    case ModuleStatus.Success: return COLORS.success;
    case ModuleStatus.Error:   return COLORS.error;
    case ModuleStatus.Running: return COLORS.warning;
    default:                   return COLORS.mutedText;
  }
}

// ── 슬라이드 헤더 공통 함수 ────────────────────────────────────
function addSlideHeader(
  slide: PptxGenJS.Slide,
  title: string,
  subtitle?: string
) {
  // 상단 배경 바
  slide.addShape("rect", {
    x: 0, y: 0, w: "100%", h: 1.1,
    fill: { color: COLORS.primary },
  });
  slide.addText(title, {
    x: 0.4, y: 0.12, w: 9.2, h: 0.6,
    fontSize: 22, bold: true, color: COLORS.white, fontFace: "맑은 고딕",
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.4, y: 0.7, w: 9.2, h: 0.35,
      fontSize: 12, color: "B0C4DE", fontFace: "맑은 고딕",
    });
  }
}

// ── 표지 슬라이드 ─────────────────────────────────────────────
function addCoverSlide(pptx: PptxGenJS, projectName: string, moduleCount: number, date: string) {
  const slide = pptx.addSlide();

  slide.addShape("rect", { x: 0, y: 0, w: "100%", h: "100%", fill: { color: COLORS.primary } });
  slide.addShape("rect", { x: 0, y: 4.5, w: "100%", h: 3, fill: { color: "163172" } });

  slide.addText("ML 분석 보고서", {
    x: 0.5, y: 1.0, w: 9, h: 0.9,
    fontSize: 36, bold: true, color: COLORS.white, fontFace: "맑은 고딕", align: "center",
  });
  slide.addText(projectName, {
    x: 0.5, y: 2.0, w: 9, h: 0.7,
    fontSize: 24, color: "93C5FD", fontFace: "맑은 고딕", align: "center",
  });
  slide.addText(`총 ${moduleCount}개 모듈 분석`, {
    x: 0.5, y: 3.0, w: 9, h: 0.5,
    fontSize: 16, color: "BFDBFE", fontFace: "맑은 고딕", align: "center",
  });
  slide.addText(date, {
    x: 0.5, y: 6.5, w: 9, h: 0.5,
    fontSize: 13, color: "93C5FD", fontFace: "맑은 고딕", align: "center",
  });
}

// ── 목차 슬라이드 ─────────────────────────────────────────────
function addTocSlide(pptx: PptxGenJS, modules: CanvasModule[]) {
  const slide = pptx.addSlide();
  addSlideHeader(slide, "목차", "분석 파이프라인 구성 모듈");

  const analysisModules = modules.filter(
    (m) => m.type !== ModuleType.TextBox && m.type !== ModuleType.GroupBox
  );

  const rows: PptxGenJS.TableRow[] = [
    [
      { text: "#", options: { bold: true, fill: { color: COLORS.primary }, color: COLORS.white, align: "center" } },
      { text: "모듈명", options: { bold: true, fill: { color: COLORS.primary }, color: COLORS.white } },
      { text: "유형", options: { bold: true, fill: { color: COLORS.primary }, color: COLORS.white } },
      { text: "상태", options: { bold: true, fill: { color: COLORS.primary }, color: COLORS.white, align: "center" } },
    ],
  ];

  analysisModules.forEach((m, i) => {
    rows.push([
      { text: String(i + 1), options: { align: "center", color: COLORS.mutedText } },
      { text: m.name },
      { text: getModuleLabel(m.type), options: { color: COLORS.mutedText } },
      { text: getStatusLabel(m.status), options: { align: "center", color: getStatusColor(m.status), bold: true } },
    ]);
  });

  (slide as any).addTable(rows, {
    x: 0.4, y: 1.3, w: 9.2,
    fontSize: 11, fontFace: "맑은 고딕",
    border: { pt: 1, color: COLORS.border },
    rowH: 0.35,
    colW: [0.6, 3.5, 3.0, 1.6],
  });
}

// ── 모듈 상세 슬라이드 ────────────────────────────────────────
function addModuleSlide(pptx: PptxGenJS, module: CanvasModule, index: number) {
  const slide = pptx.addSlide();
  addSlideHeader(
    slide,
    `${index}. ${module.name}`,
    `유형: ${getModuleLabel(module.type)}  |  상태: ${getStatusLabel(module.status)}`
  );

  let y = 1.3;

  // 파라미터 섹션
  const params = Object.entries(module.parameters ?? {});
  if (params.length > 0) {
    slide.addText("파라미터", {
      x: 0.4, y, w: 9.2, h: 0.3,
      fontSize: 13, bold: true, color: COLORS.darkText, fontFace: "맑은 고딕",
    });
    y += 0.35;

    const paramRows: PptxGenJS.TableRow[] = [
      [
        { text: "파라미터", options: { bold: true, fill: { color: "EFF6FF" }, color: COLORS.darkText } },
        { text: "값", options: { bold: true, fill: { color: "EFF6FF" }, color: COLORS.darkText } },
      ],
    ];
    params.slice(0, 8).forEach(([k, v]) => {
      const valStr = v === null || v === undefined ? "-" : String(v).slice(0, 80);
      paramRows.push([{ text: k }, { text: valStr, options: { color: COLORS.mutedText } }]);
    });

    (slide as any).addTable(paramRows, {
      x: 0.4, y, w: 9.2,
      fontSize: 10, fontFace: "맑은 고딕",
      border: { pt: 1, color: COLORS.border },
      rowH: 0.32,
      colW: [3.5, 5.7],
    });
    y += 0.32 * (paramRows.length) + 0.25;
  }

  // 출력 데이터 요약
  if (module.outputData && y < 6.5) {
    addOutputSummary(slide, module, y);
  }
}

// ── 출력 데이터 요약 ──────────────────────────────────────────
function addOutputSummary(slide: PptxGenJS.Slide, module: CanvasModule, startY: number) {
  const output = module.outputData as any;
  if (!output) return;

  slide.addText("출력 데이터 요약", {
    x: 0.4, y: startY, w: 9.2, h: 0.3,
    fontSize: 13, bold: true, color: COLORS.darkText, fontFace: "맑은 고딕",
  });
  startY += 0.35;

  const lines: string[] = [];

  if (output.type === "DataPreview") {
    const rows = output.data?.length ?? 0;
    const cols = output.columns?.length ?? 0;
    lines.push(`행: ${rows}개  |  열: ${cols}개`);
    if (output.columns?.length) {
      lines.push(`컬럼: ${(output.columns as any[]).slice(0, 8).map((c: any) => c.name ?? c).join(", ")}${output.columns.length > 8 ? " ..." : ""}`);
    }
  } else if (output.type === "TrainedModelOutput") {
    if (output.metrics) {
      Object.entries(output.metrics).slice(0, 6).forEach(([k, v]) => {
        lines.push(`${k}: ${typeof v === "number" ? (v as number).toFixed(4) : v}`);
      });
    }
  } else if (output.type === "EvaluationOutput") {
    if (output.metrics) {
      Object.entries(output.metrics).slice(0, 6).forEach(([k, v]) => {
        lines.push(`${k}: ${typeof v === "number" ? (v as number).toFixed(4) : v}`);
      });
    }
  } else if (output.type === "SplitDataOutput") {
    const trainRows = output.train?.data?.length ?? 0;
    const testRows = output.test?.data?.length ?? 0;
    lines.push(`학습 데이터: ${trainRows}행  |  테스트 데이터: ${testRows}행`);
  } else {
    // 일반 키-값 출력
    Object.entries(output)
      .filter(([k]) => k !== "type" && typeof output[k] !== "object")
      .slice(0, 6)
      .forEach(([k, v]) => lines.push(`${k}: ${v}`));
  }

  if (lines.length > 0) {
    slide.addText(lines.join("\n"), {
      x: 0.4, y: startY, w: 9.2, h: Math.min(lines.length * 0.32 + 0.1, 2.0),
      fontSize: 11, color: COLORS.mutedText, fontFace: "맑은 고딕",
      fill: { color: COLORS.lightBg },
      line: { color: COLORS.border, pt: 1 },
      margin: 6,
    });
  }
}

// ── 파이프라인 요약 슬라이드 ──────────────────────────────────
function addPipelineSummarySlide(pptx: PptxGenJS, modules: CanvasModule[], connections: Connection[]) {
  const slide = pptx.addSlide();
  addSlideHeader(slide, "파이프라인 요약", "모듈 연결 구조");

  const analysisModules = modules.filter(
    (m) => m.type !== ModuleType.TextBox && m.type !== ModuleType.GroupBox
  );

  const total = analysisModules.length;
  const success = analysisModules.filter((m) => m.status === ModuleStatus.Success).length;
  const errors  = analysisModules.filter((m) => m.status === ModuleStatus.Error).length;
  const pending = analysisModules.filter((m) => m.status === ModuleStatus.Pending).length;

  const stats = [
    { label: "전체 모듈", value: total, color: COLORS.primary },
    { label: "성공", value: success, color: COLORS.success },
    { label: "오류", value: errors,  color: COLORS.error },
    { label: "대기/실행 중", value: pending, color: COLORS.mutedText },
    { label: "연결 수", value: connections.length, color: COLORS.accent },
  ];

  stats.forEach((s, i) => {
    const x = 0.4 + i * 1.9;
    slide.addShape("rect", { x, y: 1.3, w: 1.7, h: 1.0, fill: { color: COLORS.lightBg }, line: { color: COLORS.border, pt: 1 } });
    slide.addText(String(s.value), { x, y: 1.4, w: 1.7, h: 0.5, fontSize: 26, bold: true, color: s.color, fontFace: "맑은 고딕", align: "center" });
    slide.addText(s.label, { x, y: 1.95, w: 1.7, h: 0.3, fontSize: 10, color: COLORS.mutedText, fontFace: "맑은 고딕", align: "center" });
  });

  // 연결 목록
  if (connections.length > 0) {
    slide.addText("모듈 연결 목록", {
      x: 0.4, y: 2.55, w: 9.2, h: 0.3,
      fontSize: 13, bold: true, color: COLORS.darkText, fontFace: "맑은 고딕",
    });

    const connRows: PptxGenJS.TableRow[] = [
      [
        { text: "출발 모듈", options: { bold: true, fill: { color: COLORS.primary }, color: COLORS.white } },
        { text: "출발 포트", options: { bold: true, fill: { color: COLORS.primary }, color: COLORS.white } },
        { text: "도착 모듈", options: { bold: true, fill: { color: COLORS.primary }, color: COLORS.white } },
        { text: "도착 포트", options: { bold: true, fill: { color: COLORS.primary }, color: COLORS.white } },
      ],
    ];

    connections.slice(0, 12).forEach((c) => {
      const fromModule = modules.find((m) => m.id === c.from.moduleId);
      const toModule   = modules.find((m) => m.id === c.to.moduleId);
      connRows.push([
        { text: fromModule?.name ?? c.from.moduleId, options: { color: COLORS.darkText } },
        { text: c.from.portName, options: { color: COLORS.mutedText } },
        { text: toModule?.name ?? c.to.moduleId, options: { color: COLORS.darkText } },
        { text: c.to.portName, options: { color: COLORS.mutedText } },
      ]);
    });

    (slide as any).addTable(connRows, {
      x: 0.4, y: 2.9, w: 9.2,
      fontSize: 10, fontFace: "맑은 고딕",
      border: { pt: 1, color: COLORS.border },
      rowH: 0.32,
      colW: [2.8, 1.8, 2.8, 1.8],
    });
  }
}

// ── 메인 빌더 함수 ─────────────────────────────────────────────
export async function buildSlideReport(
  modules: CanvasModule[],
  connections: Connection[],
  projectName: string
): Promise<void> {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE"; // 16:9

  const today = new Date().toLocaleDateString("ko-KR", {
    year: "numeric", month: "long", day: "numeric",
  });

  const analysisModules = modules.filter(
    (m) => m.type !== ModuleType.TextBox && m.type !== ModuleType.GroupBox
  );

  // 1. 표지
  addCoverSlide(pptx, projectName, analysisModules.length, today);

  // 2. 목차
  if (analysisModules.length > 0) {
    addTocSlide(pptx, analysisModules);
  }

  // 3. 파이프라인 요약
  addPipelineSummarySlide(pptx, modules, connections);

  // 4. 각 모듈 상세 슬라이드
  analysisModules.forEach((m, i) => {
    addModuleSlide(pptx, m, i + 1);
  });

  // 저장
  const fileName = `${projectName}_ML분석보고서`;
  await pptx.writeFile({ fileName });
}
